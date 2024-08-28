using System.Globalization;
using System.IO.Pipes;
using System.Text;
using Windows.Devices.Bluetooth;

namespace BLE_EmulatorTest;

class Program
{
    private static VirtualKeyboard m_virtualKeyboard;
    private static VirtualMouse m_virtualMouse;

    private static async void InitializeVirtualDevices()
    {
        try
        {
            m_virtualKeyboard = new VirtualKeyboard();
            m_virtualKeyboard.SubscribedHidClientsChanged += VirtualKeyboard_SubscribedHidClientsChanged;
            await m_virtualKeyboard.InitilizeAsync();
            m_virtualKeyboard.Enable();

            m_virtualMouse = new VirtualMouse();
            m_virtualMouse.SubscribedHidClientsChanged += VirtualMouse_SubscribedHidClientsChanged;
            await m_virtualMouse.InitilizeAsync();
            m_virtualMouse.Enable();
        }
        catch (Exception e)
        {
            Console.WriteLine("Error: " + e.ToString());
        }

        //await Test();
        await run_server();
    }

    private static async void VirtualKeyboard_SubscribedHidClientsChanged(IReadOnlyList<Windows.Devices.Bluetooth.GenericAttributeProfile.GattSubscribedClient> subscribedClients)
    {
        if (subscribedClients != null)
        {
            foreach (var client in subscribedClients)
            {
                var leDevice = await BluetoothLEDevice.FromIdAsync(client.Session.DeviceId.Id);
                Console.WriteLine("keyboard-subscribed: " + leDevice.Name);
            }
        }
    }

    private static async void VirtualMouse_SubscribedHidClientsChanged(IReadOnlyList<Windows.Devices.Bluetooth.GenericAttributeProfile.GattSubscribedClient> subscribedClients)
    {
        if (subscribedClients != null)
        {
            foreach (var client in subscribedClients)
            {
                var leDevice = await BluetoothLEDevice.FromIdAsync(client.Session.DeviceId.Id);
                Console.WriteLine("mouse-subscribed: " + leDevice.Name);
            }
        }
    }

    private static async Task Test()
    {
        await Task.Delay(1000);

        for (int i = 0; i < 10; i++)
        {
            await m_virtualKeyboard.PressKey(0x05);
            await Task.Delay(100);
            await m_virtualKeyboard.ReleaseKey(0x05);
            await Task.Delay(100);
            await m_virtualKeyboard.PressKey(0x1e);
            await Task.Delay(100);
            await m_virtualKeyboard.ReleaseKey(0x1e);
            await Task.Delay(100);
        }

        for (int i = 0; i < 10; i++)
        {
            await m_virtualMouse.Move(-10, -10, 0);
            await Task.Delay(300);
        }
    
        for (int i = 0; i < 10; i++)
        {
            await m_virtualMouse.Move(10, 10, 0);
            await Task.Delay(300);
        }

        for (int i = 0; i < 4; i++)
        {
            m_virtualMouse.Press();
            await Task.Delay(300);
            m_virtualMouse.Release();
            await Task.Delay(300);
        }
    }

    private static BinaryReader br = null;
    private static BinaryWriter bw = null;

    private static string ReadString()
    {
        var len = (int)br.ReadUInt32();            // Read string length
        var str = new string(br.ReadChars(len));    // Read string
        Console.WriteLine("Read: \"{0}\"", str);
        return str;
    }

    private static void WriteString(string str)
    {
        var buf = Encoding.ASCII.GetBytes(str);     // Get ASCII byte array     
        bw.Write((uint)buf.Length);                // Write string length
        bw.Write(buf);                              // Write string
        Console.WriteLine("Wrote: \"{0}\"", str);
    }

    private static async Task run_server()
    {
        // Open the named pipe.
        NamedPipeServerStream server = null;
        bool waitingForConnection = true;

        while (true)
        {
            if (waitingForConnection)
            {
                Console.WriteLine("Waiting for connection...");
                server = new NamedPipeServerStream("BV_BLE_PIPE");
                server.WaitForConnection();

                Console.WriteLine("Connected.");
                br = new BinaryReader(server);
                bw = new BinaryWriter(server);
                waitingForConnection = false;
            }

            try
            {
                var str = ReadString();

                if (str == "AT+BLECURRENTDEVICENAME\r\n")
                {
                    WriteString("at + blecurrentdevicename\n");
                    WriteString("OK\n");
                }
                else if (str.StartsWith("AT+BLEHIDMOUSEMOVE"))
                {
                    var args = str.Split(new char[] { '=', '\r', '\n' })[1].Split(',');
                    await m_virtualMouse.Move(int.Parse(args[0]), int.Parse(args[1]), int.Parse(args[2]));
                    WriteString("OK\n");
                }
                else if (str.StartsWith("AT+BLEHIDMOUSEBUTTON"))
                {
                    var args = str.Split(new char[] { '=', '\r', '\n' })[1].Split(',');
                    if (args[0] == "l")
                    {
                        if (args[1] == "press")
                        {
                            await m_virtualMouse.Press();
                        }
                        else if (args[1] == "click")
                        {
                            await m_virtualMouse.Click();
                        }
                    }
                    else if (args[0] == "0")
                    {
                        await m_virtualMouse.Release();
                    }
                    WriteString("OK\n");
                }
                else if (str.StartsWith("AT+BLEKEYBOARDCODE"))
                {
                    var args = str.Split(new char[] { '=', '\r', '\n' })[1].Split('-');
                    var reportValue = new byte[VirtualKeyboard.c_sizeOfKeyboardReportDataInBytes];
                    for (int i = 0; i < args.Length; i++)
                    {
                        reportValue[i] = byte.Parse(args[i], NumberStyles.HexNumber);
                    }
                    await m_virtualKeyboard.DirectSendReport(reportValue);
                    WriteString("OK\n");
                }
                else
                {
                    Console.WriteLine("UNKNOWN CMD!!!");
                    WriteString("OK\n"); //always OK for now
                }
            }
            // When client disconnects
            catch (EndOfStreamException)
            {
                Console.WriteLine("Client disconnected.");
                server.Close();
                server.Dispose();
                waitingForConnection = true;
            }
        }
    }

    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        InitializeVirtualDevices();
        //Console.ReadLine();
        while (true) ;
    }
}