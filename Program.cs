using BluetoothLEExplorer.Models;
using Windows.Devices.Bluetooth;

namespace BLE_EmulatorTest;

class Program
{
    private static VirtualKeyboard m_virtualKeyboard;

    private static async void InitializeVirtualKeyboard()
    {
        try
        {
            m_virtualKeyboard = new VirtualKeyboard();
            m_virtualKeyboard.SubscribedHidClientsChanged += VirtualKeyboard_SubscribedHidClientsChanged;
            await m_virtualKeyboard.InitiliazeAsync();
            m_virtualKeyboard.Enable();
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
                Console.WriteLine("subscribed: " + leDevice.Name);
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
    }

    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        InitializeVirtualKeyboard();
        Console.ReadLine();
    }
}
