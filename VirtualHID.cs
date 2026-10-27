using System.Diagnostics;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Storage.Streams;

namespace BLE_EmulatorTest;
class VirtualHID
{
    private static readonly GattLocalCharacteristicParameters c_hidInputReportParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Notify,
        ReadProtectionLevel = GattProtectionLevel.EncryptionRequired,
        WriteProtectionLevel = GattProtectionLevel.EncryptionRequired
    };

    private static readonly GattLocalCharacteristicParameters c_hidReportMapParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read,
        ReadProtectionLevel = GattProtectionLevel.EncryptionRequired,
        WriteProtectionLevel = GattProtectionLevel.EncryptionRequired,
        StaticValue = new byte[]
        {
            0x05, 0x01,        // Usage Page (Generic Desktop Ctrls)
            0x09, 0x06,        // Usage (Keyboard)
            0xA1, 0x01,        // Collection (Application)
            0x85, 0x01,        //   Report ID
            0x05, 0x07,        //   Usage Page (Kbrd/Keypad)
            0x19, 0xE0,        //   Usage Minimum (0xE0)
            0x29, 0xE7,        //   Usage Maximum (0xE7)
            0x15, 0x00,        //   Logical Minimum (0)
            0x25, 0x01,        //   Logical Maximum (1)
            0x95, 0x08,        //   Report Count (8)
            0x75, 0x01,        //   Report Size (1)
            0x81, 0x02,        //   Input (Data,Var,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0x95, 0x01,        //   Report Count (1)
            0x75, 0x08,        //   Report Size (8)
            0x81, 0x01,        //   Input (Const,Array,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0x05, 0x07,        //   Usage Page (Kbrd/Keypad)
            0x19, 0x00,        //   Usage Minimum (0x00)
            0x2a, 0xff, 0x00,  //   Usage Maximum (255)
            0x15, 0x00,        //   Logical Minimum (0)
            0x26, 0xff, 0x00,  //   Logical Maximum (255)
            0x95, 0x06,        //   Report Count (6)
            0x75, 0x08,        //   Report Size (8)
            0x81, 0x00,        //   Input (Data,Array,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0xC0,              // End Collection


            0x05, 0x01,        // Usage Page (Generic Desktop Ctrls)
            0x09, 0x02,        // Usage (Mouse)
            0xA1, 0x01,        // Collection (Application)
            0x85, 0x02,        //   Report ID (2)
            0x09, 0x01,        //   Usage (Pointer)
            0xA1, 0x00,        //   Collection (Physical)
            0x05, 0x09,        //     Usage Page (Button)
            0x19, 0x01,        //     Usage Minimum (0x01)
            0x29, 0x02,        //     Usage Maximum (0x02)
            0x15, 0x00,        //     Logical Minimum (0)
            0x25, 0x01,        //     Logical Maximum (1)
            0x75, 0x01,        //     Report Size (1)
            0x95, 0x02,        //     Report Count (2)
            0x81, 0x02,        //     Input (Data,Var,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0x95, 0x06,        //     Report Count (6)
            0x81, 0x03,        //     Input (Const,Var,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0x05, 0x01,        //     Usage Page (Generic Desktop Ctrls)
            0x09, 0x30,        //     Usage (X)
            0x09, 0x31,        //     Usage (Y)
            0x09, 0x38,        //     Usage (Wheel)
            0x16, 0x01, 0x80,  //     Logical Minimum (-32767)
            0x26, 0xFF, 0x7F,  //     Logical Maximum (32767)
            0x75, 0x10,        //     Report Size (16)
            0x95, 0x03,        //     Report Count (3)
            0x81, 0x06,        //     Input (Data,Var,Rel,No Wrap,Linear,Preferred State,No Null Position)
            0xC0,              //   End Collection
            0xC0,              // End Collection

            0x05, 0x0C,        // Usage Page (Consumer)
            0x09, 0x01,        // Usage (Consumer Control)
            0xA1, 0x01,        // Collection (Application)
            0x85, 0x03,        //   Report ID (3)
            0x05, 0x0C,        //   Usage Page (Consumer)
            0x15, 0x00,        //   Logical Minimum (0)
            0x26, 0xff, 0x03,  //   Logical Maximum (0x3FF)
            0x19, 0x00,        //   Usage Minimum (0x00)
            0x2a, 0xff, 0x03,  //   Usage Maximum (0x3FF)
            0x75, 0x10,        //   Report Size (16)
            0x95, 0x01,        //   Report Count (1)
            0x81, 0x00,        //   Input (Data,Array,Abs,No Wrap,Linear,Preferred State,No Null Position)
            0xC0,              // End Collection

        }.AsBuffer()
    };

    private static readonly GattLocalCharacteristicParameters c_hidInformationParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read,
        ReadProtectionLevel = GattProtectionLevel.Plain,
        StaticValue = new byte[]
        {
            0x11, 0x01, // HID Version: 1101
            0x00,       // Country Code: 0
            0x01        // Not Normally Connectable, Remote Wake supported
        }.AsBuffer()
    };

    private static readonly GattLocalCharacteristicParameters c_hidControlPointParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.WriteWithoutResponse,
        WriteProtectionLevel = GattProtectionLevel.Plain
    };

    public static readonly uint c_mouseReportSize = 0x7;
    public static readonly uint c_keyboardReportSize = 0x8;
    public static readonly uint c_consumerReportSize = 0x2;

    private GattServiceProvider m_hidServiceProvider;
    private GattLocalService m_hidService;
    private GattLocalCharacteristic m_keyboardReport, m_consumerReport, m_mouseReport;
    private GattLocalCharacteristic m_hidReportMap;
    private GattLocalCharacteristic m_hidInformation;
    private GattLocalCharacteristic m_hidControlPoint;

    private Object m_lock = new Object();

    private bool m_initializationFinished = false;

    private bool m_lastLeftDown = false;
    private bool m_lastRightDown = false;

    public delegate void SubscribedHidClientsChangedHandler(IReadOnlyList<GattSubscribedClient> subscribedClients);

    private static string GetStringFromBuffer(IBuffer buffer)
    {
        return GetStringFromBuffer(buffer.ToArray());
    }

    private static string GetStringFromBuffer(byte[] bytes)
    {
        return BitConverter.ToString(bytes).Replace("-", " ");
    }

    public async Task InitilizeAsync()
    {
        await CreateHidService();

        lock (m_lock)
        {
            m_initializationFinished = true;
        }
    }

    public void Enable()
    {
        PublishService(m_hidServiceProvider);
    }

    public void Disable()
    {
        UnpublishService(m_hidServiceProvider);
    }

    private async Task<GattLocalCharacteristic> setupReport(byte reportID)
    {
        // HID Report characteristic.
        var reportResult = await m_hidService.CreateCharacteristicAsync(GattCharacteristicUuids.Report, c_hidInputReportParameters);
        if (reportResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the report characteristic: " + reportResult.Error);
            throw new Exception("Failed to create the report characteristic: " + reportResult.Error);
        }
        GattLocalCharacteristic report = reportResult.Characteristic;

        GattLocalDescriptorParameters referenceParams = new GattLocalDescriptorParameters
        {
            ReadProtectionLevel = GattProtectionLevel.EncryptionRequired,
            WriteProtectionLevel = GattProtectionLevel.EncryptionRequired,
            StaticValue = new byte[] { reportID, 0x01 }.AsBuffer()
        };

        // HID Report Reference descriptor.
        var referenceResult = await report.CreateDescriptorAsync(BluetoothUuidHelper.FromShortId(0x2908), referenceParams);
        if (referenceResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the report reference descriptor: " + referenceResult.Error);
            throw new Exception("Failed to create the report reference descriptor: " + referenceResult.Error);
        }

        report.SubscribedClientsChanged += SubscribedClientsChanged;
        return report;
    }

    private async Task CreateHidService()
    {
        // HID service.
        var providerResult = await GattServiceProvider.CreateAsync(GattServiceUuids.HumanInterfaceDevice);
        if (providerResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the HID service provider: " + providerResult.Error);
            throw new Exception("Failed to create the HID service provider: " + providerResult.Error);
        }
        m_hidServiceProvider = providerResult.ServiceProvider;
        m_hidService = m_hidServiceProvider.Service;

        m_keyboardReport = await setupReport(1);
        m_mouseReport = await setupReport(2);
        m_consumerReport = await setupReport(3);

        // HID Report Map characteristic.
        var hidReportMapCharacteristicCreationResult = await m_hidService.CreateCharacteristicAsync(GattCharacteristicUuids.ReportMap, c_hidReportMapParameters);
        if (hidReportMapCharacteristicCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the HID report map characteristic: " + hidReportMapCharacteristicCreationResult.Error);
            throw new Exception("Failed to create the HID report map characteristic: " + hidReportMapCharacteristicCreationResult.Error);
        }
        m_hidReportMap = hidReportMapCharacteristicCreationResult.Characteristic;

        // HID Information characteristic.
        var hidInformationCharacteristicCreationResult = await m_hidService.CreateCharacteristicAsync(GattCharacteristicUuids.HidInformation, c_hidInformationParameters);
        if (hidInformationCharacteristicCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the HID information characteristic: " + hidInformationCharacteristicCreationResult.Error);
            throw new Exception("Failed to create the HID information characteristic: " + hidInformationCharacteristicCreationResult.Error);
        }
        m_hidInformation = hidInformationCharacteristicCreationResult.Characteristic;

        // HID Control Point characteristic.
        var hidControlPointCharacteristicCreationResult = await m_hidService.CreateCharacteristicAsync(GattCharacteristicUuids.HidControlPoint, c_hidControlPointParameters);
        if (hidControlPointCharacteristicCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the HID control point characteristic: " + hidControlPointCharacteristicCreationResult.Error);
            throw new Exception("Failed to create the HID control point characteristic: " + hidControlPointCharacteristicCreationResult.Error);
        }
        m_hidControlPoint = hidControlPointCharacteristicCreationResult.Characteristic;
        m_hidControlPoint.WriteRequested += HidControlPoint_WriteRequested;

        m_hidServiceProvider.AdvertisementStatusChanged += HidServiceProvider_AdvertisementStatusChanged;
    }

    // Assumes that the lock is being held.
    private void PublishService(GattServiceProvider provider)
    {
        var advertisingParameters = new GattServiceProviderAdvertisingParameters
        {
            IsDiscoverable = true,
            IsConnectable = true // Peripheral role support is required for Windows to advertise as connectable.
        };

        provider.StartAdvertising(advertisingParameters);
    }

    private void UnpublishService(GattServiceProvider provider)
    {
        try
        {
            if ((provider.AdvertisementStatus == GattServiceProviderAdvertisementStatus.Started) ||
                (provider.AdvertisementStatus == GattServiceProviderAdvertisementStatus.Aborted))
            {
                provider.StopAdvertising();
            }
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to stop advertising due to: " + e.Message);
        }
    }

    private async void HidControlPoint_WriteRequested(GattLocalCharacteristic sender, GattWriteRequestedEventArgs args)
    {
        try
        {
            var deferral = args.GetDeferral();
            var writeRequest = await args.GetRequestAsync();
            Debug.WriteLine("Value written to HID Control Point: " + GetStringFromBuffer(writeRequest.Value));
            // Control point only supports WriteWithoutResponse.
            deferral.Complete();
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to handle write to Hid Control Point due to: " + e.Message);
        }
    }

    private void HidServiceProvider_AdvertisementStatusChanged(GattServiceProvider sender, GattServiceProviderAdvertisementStatusChangedEventArgs args)
    {
        Debug.WriteLine("HID advertisement status changed to " + args.Status);
    }

    private async void SubscribedClientsChanged(GattLocalCharacteristic sender, object args)
    {
        try
        {
            if (sender.SubscribedClients != null)
            {
                foreach (var client in sender.SubscribedClients)
                {
                    var leDevice = await BluetoothLEDevice.FromIdAsync(client.Session.DeviceId.Id);

                    // this one we'll use to actually track connect/disconnect
                    if (sender == m_mouseReport)
                    {
                        Program.LogInfo($"mouse subscription: {leDevice.DeviceId} - {leDevice.Name}");
                        await Program.deviceSubscribed(leDevice);
                    }
                    else if (sender == m_keyboardReport)
                    {
                        Program.LogInfo($"keyboard subscription: {leDevice.DeviceId} - {leDevice.Name}");
                    }
                    else if (sender == m_consumerReport)
                    {
                        Program.LogInfo($"consumer subscription: {leDevice.DeviceId} - {leDevice.Name}");
                    }
                }
            }
        }
        catch (Exception e)
        {
            Program.LogInfo($"ERROR=SubscribedClientsChanged exception: {e}");
        }
    }

    public async Task Move(short mx, short my, short wheel)
    {
        try
        {
            await SendMouseState(m_lastLeftDown, m_lastRightDown, mx, my, wheel);
            await Task.Delay(40);
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to change the mouse state due to: " + e.Message);
        }
    }

    public async Task Press()
    {
        try
        {
            await SendMouseState(true, false, 0, 0, 0);
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to change the mouse state due to: " + e.Message);
        }
    }

    public async Task Release()
    {
        try
        {
            await SendMouseState(false, false, 0, 0, 0);
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to change the mouse state due to: " + e.Message);
        }
    }

    private async Task SendMouseState(bool leftDown, bool rightDown, short mx, short my, short wheel)
    {
        //lock (m_lock)
        {
            if (!m_initializationFinished)
            {
                return;
            }

            if (m_mouseReport.SubscribedClients.Count == 0)
            {
                Debug.WriteLine("No clients are currently subscribed to the mouse report.");
                return;
            }

            var reportValue = new byte[c_mouseReportSize];

            // The first byte of the report data is buttons bitfield.
            reportValue[0] = (byte)((leftDown ? (1 << 0) : 0) | (rightDown ? (1 << 1) : 0));

            reportValue[1] = (byte)(mx & 0xFF);
            reportValue[2] = (byte)((mx >> 8) & 0xFF);

            reportValue[3] = (byte)(my & 0xFF);
            reportValue[4] = (byte)((my >> 8) & 0xFF);
            
            reportValue[5] = (byte)(wheel & 0xFF);
            reportValue[6] = (byte)((wheel >> 8) & 0xFF);

            Debug.WriteLine("Sending mouse report value notification with data: " + GetStringFromBuffer(reportValue));
            m_lastLeftDown = leftDown;
            m_lastRightDown = rightDown;

            // we'll wait for op to complete, but not necessary
            await m_mouseReport.NotifyValueAsync(reportValue.AsBuffer());
        }
    }

    public async Task SendKeyboardReport(byte[] reportValue)
    {
        if (reportValue.Length != c_keyboardReportSize)
        {
            Console.WriteLine("wrong keyboard report size!");
            return;
        }

        await m_keyboardReport.NotifyValueAsync(reportValue.AsBuffer());
    }

    public async Task SendConsumerReport(byte[] reportValue)
    {
        if (reportValue.Length != c_consumerReportSize)
        {
            Console.WriteLine("wrong consumer report size!");
            return;
        }

        await m_consumerReport.NotifyValueAsync(reportValue.AsBuffer());
    }
}
