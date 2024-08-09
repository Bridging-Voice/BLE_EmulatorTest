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

        Test();
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

    private static void Test()
    {
        Thread.Sleep(1000);
        for (int i = 0; i < 10; i++)
        {
            m_virtualKeyboard.PressKey(0x05);
            Thread.Sleep(100);
            m_virtualKeyboard.ReleaseKey(0x05);
            Thread.Sleep(100);
            m_virtualKeyboard.PressKey(0x1e);
            Thread.Sleep(100);
            m_virtualKeyboard.ReleaseKey(0x1e);
            Thread.Sleep(100);
        }

        for (int i = 0; i < 10; i++)
        {
            m_virtualMouse.MoveMouse(-10, -10);
            Thread.Sleep(300);
        }
    
        for (int i = 0; i < 10; i++)
        {
            m_virtualMouse.MoveMouse(10, 10);
            Thread.Sleep(300);
        }
    }

    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        InitializeVirtualDevices();
        Console.ReadLine();
    }
}