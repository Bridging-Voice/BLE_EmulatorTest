using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Pipes;
using System.Text;
using Windows.Devices.Bluetooth;

namespace BLE_EmulatorTest;

class Program
{
    private static VirtualKeyboard? m_virtualKeyboard;
    private static VirtualMouse? m_virtualMouse;
    private static string m_deviceName = "NONE";
    private static string VER = "BVEM 1.0.1";
    private static NamedPipeServerStream? m_pipeIn, m_pipeOut;
    private static BlockingCollection<string> m_cmds = new BlockingCollection<string>();
    private static Thread? m_readThread;

    private static async Task InitializeVirtualDevices()
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
                m_deviceName = leDevice.Name;
                WriteString("DEVICE=" + m_deviceName + "\n");
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
                m_deviceName = leDevice.Name;
                WriteString("DEVICE=" + m_deviceName + "\n");
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
            await m_virtualMouse.Press();
            await Task.Delay(300);
            await m_virtualMouse.Release();
            await Task.Delay(300);
        }
    }

    private static void ReadStrings()
    {
        while (true)
        {
            try
            {
                if (m_pipeIn is null || !m_pipeIn.IsConnected)
                {
                    Thread.Sleep(500);
                }
                else
                {
                    byte[] buf = new byte[128];
                    m_pipeIn.Read(buf, 0, buf.Length);
                    var str = System.Text.Encoding.ASCII.GetString(buf);
                    m_cmds.Add(str);
                    Console.WriteLine("Read: \"{0}\"", str);
                }
            }
            catch { Thread.Sleep(500); }
        }
    }

    private static void WriteString(string str)
    {
        //Console.WriteLine("entering WriteString");
        var buf = Encoding.ASCII.GetBytes(str);     // Get ASCII byte array
        if (m_pipeOut is not null && m_pipeOut.IsConnected)
        {
            m_pipeOut.Write(buf);
        }
        Console.WriteLine("Wrote: \"{0}\"", str);
    }

    private static async Task run_server()
    {
        m_readThread = new Thread(ReadStrings);
        m_readThread.Start();

        // Open the named pipe.
        bool waitingForConnection = true;

        while (true)
        {
            if (waitingForConnection)
            {
                Console.WriteLine("Waiting for connection...");
                m_pipeIn = new NamedPipeServerStream("BV_BLE_PIPE_IN", PipeDirection.InOut, 1, PipeTransmissionMode.Message, PipeOptions.WriteThrough, 0, 0);
                m_pipeOut = new NamedPipeServerStream("BV_BLE_PIPE_OUT", PipeDirection.InOut, 1, PipeTransmissionMode.Message, PipeOptions.WriteThrough, 0, 0);
                m_pipeIn.WaitForConnection();
                m_pipeOut.WaitForConnection();

                Console.WriteLine("Connected.");
                waitingForConnection = false;
            }

            try
            {
                var str = m_cmds.Take().Replace("\n", "").Replace("\r", "").TrimEnd('\0').ToLower();
                WriteString("\x06\n"); // ACK when we get to actually processing it

                if (str.StartsWith("begin"))
                {
                    WriteString(VER + "\n");
                    WriteString("DEVICE=" + m_deviceName + "\n");
                }
                else if (str.StartsWith("ping"))
                {
                    WriteString("pong\n");
                }
                else if (str.StartsWith("mv"))
                {
                    var args = str.Split(new char[] { '=' })[1].Split(',');
                    await m_virtualMouse.Move(int.Parse(args[0]), int.Parse(args[1]), 0);
                    WriteString("OK\n");
                }
                else if (str.StartsWith("mw"))
                {
                    var args = str.Split(new char[] { '=' })[1].Split(',');
                    await m_virtualMouse.Move(0, 0, int.Parse(args[0]));
                    WriteString("OK\n");
                }
                else if (str.StartsWith("mb"))
                {
                    var args = str.Split(new char[] { '=' })[1].Split(',');
                    if (args[0] == "l")
                    {
                        if (args[1] == "1")
                        {
                            await m_virtualMouse.Press();
                        }
                        else if (args[1] == "0")
                        {
                            await m_virtualMouse.Release();
                        }
                    }
                    WriteString("OK\n");
                }
                else if (str.StartsWith("kb"))
                {
                    var args = str.Split(new char[] { '=' })[1].Split('-');
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
                    WriteString("ERROR=unknown cmd\n");
                }
            }
            // When client disconnects
            catch (System.IO.IOException)
            {
                Console.WriteLine("Client disconnected.");
                if (m_pipeIn is not null)
                {
                    m_pipeIn.Close();
                    m_pipeIn.Dispose();
                    m_pipeIn = null;
                }
                if (m_pipeOut is not null)
                {
                    m_pipeOut.Close();
                    m_pipeOut.Dispose();
                    m_pipeOut = null;
                }
                waitingForConnection = true;
            }
        }
    }

    static async Task Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        await InitializeVirtualDevices();
    }
}