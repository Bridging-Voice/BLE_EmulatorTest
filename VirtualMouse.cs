using System.Runtime.InteropServices.WindowsRuntime;
using System.Diagnostics;

using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Storage.Streams;

namespace BLE_EmulatorTest;
class VirtualMouse
{
    private static readonly GattLocalCharacteristicParameters c_hidInputReportParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Notify,
        ReadProtectionLevel = GattProtectionLevel.EncryptionRequired
    };

    private static readonly uint c_hidReportReferenceDescriptorShortUuid = 0x2908;

    private static readonly GattLocalDescriptorParameters c_hidMouseReportReferenceParameters = new GattLocalDescriptorParameters
    {
        ReadProtectionLevel = GattProtectionLevel.EncryptionRequired,
        StaticValue = new byte[]
        {
            0x02, // Report ID: 1
            0x01  // Report Type: Input
        }.AsBuffer()
    };

    private static readonly GattLocalCharacteristicParameters c_hidReportMapParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read,
        ReadProtectionLevel = GattProtectionLevel.EncryptionRequired,                        
        StaticValue = new byte[]
        {
            0x05, 0x01,                         // USAGE_PAGE (Generic Desktop)     0
            0x09, 0x02,                         // USAGE (Mouse)                    2
            0xa1, 0x01,                         // COLLECTION (Application)         4
            0x85, 0x02,                         //   REPORT_ID (Mouse)              6
            0x09, 0x01,                         //   USAGE (Pointer)                8
            0xa1, 0x00,                         //   COLLECTION (Physical)          10
            0x05, 0x09,                         //     USAGE_PAGE (Button)          12
            0x19, 0x01,                         //     USAGE_MINIMUM (Button 1)     14
            0x29, 0x02,                         //     USAGE_MAXIMUM (Button 2)     16
            0x15, 0x00,                         //     LOGICAL_MINIMUM (0)          18
            0x25, 0x01,                         //     LOGICAL_MAXIMUM (1)          20
            0x75, 0x01,                         //     REPORT_SIZE (1)              22
            0x95, 0x02,                         //     REPORT_COUNT (2)             24
            0x81, 0x02,                         //     INPUT (Data,Var,Abs)         26
            0x95, 0x06,                         //     REPORT_COUNT (6)             28
            0x81, 0x03,                         //     INPUT (Cnst,Var,Abs)         30
            0x05, 0x01,                         //     USAGE_PAGE (Generic Desktop) 32
            0x09, 0x30,                         //     USAGE (X)                    34
            0x09, 0x31,                         //     USAGE (Y)                    36
            0x15, 0x81,                         //     LOGICAL_MINIMUM (-127)       38
            0x25, 0x7f,                         //     LOGICAL_MAXIMUM (127)        40
            0x75, 0x08,                         //     REPORT_SIZE (8)              42
            0x95, 0x02,                         //     REPORT_COUNT (2)             44
            0x81, 0x06,                         //     INPUT (Data,Var,Rel)         46
            0xc0,                               //   END_COLLECTION                 48
            0xc0                                // END_COLLECTION                   49/50
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

    private static readonly GattLocalCharacteristicParameters c_batteryLevelParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Notify,
        ReadProtectionLevel = GattProtectionLevel.Plain
    };

    private static readonly uint c_sizeOfMouseReportDataInBytes = 0x3;

    private GattServiceProvider m_hidServiceProvider;
    private GattLocalService m_hidService;
    private GattLocalCharacteristic m_hidMouseReport;
    private GattLocalDescriptor m_hidMouseReportReference;
    private GattLocalCharacteristic m_hidReportMap;
    private GattLocalCharacteristic m_hidInformation;
    private GattLocalCharacteristic m_hidControlPoint;

    private Object m_lock = new Object();

    private bool m_initializationFinished = false;

    private bool m_lastLeftDown = false;
    private bool m_lastRightDown = false;

    public delegate void SubscribedHidClientsChangedHandler(IReadOnlyList<GattSubscribedClient> subscribedClients);
    public event SubscribedHidClientsChangedHandler SubscribedHidClientsChanged;

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

    public async Task Move(int mx, int my)
    {
        try
        {
            await SendMouseState(m_lastLeftDown, m_lastRightDown, mx, my);
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
            await SendMouseState(true, false, 0, 0);
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
            await SendMouseState(false, false, 0, 0);
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to change the mouse state due to: " + e.Message);
        }
    }

    public async Task Click()
    {
        try
        {
            await SendMouseState(true, false, 0, 0);
            await Task.Delay(40);
            await SendMouseState(false, false, 0, 0);
        }
        catch (Exception e)
        {
            Debug.WriteLine("Failed to change the mouse state due to: " + e.Message);
        }
    }

    private async Task CreateHidService()
    {
        // HID service.
        var hidServiceProviderCreationResult = await GattServiceProvider.CreateAsync(GattServiceUuids.HumanInterfaceDevice);
        if (hidServiceProviderCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the HID service provider: " + hidServiceProviderCreationResult.Error);
            throw new Exception("Failed to create the HID service provider: " + hidServiceProviderCreationResult.Error);
        }
        m_hidServiceProvider = hidServiceProviderCreationResult.ServiceProvider;
        m_hidService = m_hidServiceProvider.Service;

        // HID mouse Report characteristic.
        var hidMouseReportCharacteristicCreationResult = await m_hidService.CreateCharacteristicAsync(GattCharacteristicUuids.Report, c_hidInputReportParameters);
        if (hidMouseReportCharacteristicCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the mouse report characteristic: " + hidMouseReportCharacteristicCreationResult.Error);
            throw new Exception("Failed to create the mouse report characteristic: " + hidMouseReportCharacteristicCreationResult.Error);
        }
        m_hidMouseReport = hidMouseReportCharacteristicCreationResult.Characteristic;
        m_hidMouseReport.SubscribedClientsChanged += HidMouseReport_SubscribedClientsChanged;

        // HID mouse Report Reference descriptor.
        var hidMouseReportReferenceCreationResult = await m_hidMouseReport.CreateDescriptorAsync(BluetoothUuidHelper.FromShortId(c_hidReportReferenceDescriptorShortUuid), c_hidMouseReportReferenceParameters);
        if (hidMouseReportReferenceCreationResult.Error != BluetoothError.Success)
        {
            Debug.WriteLine("Failed to create the mouse report reference descriptor: " + hidMouseReportReferenceCreationResult.Error);
            throw new Exception("Failed to create the mouse report reference descriptor: " + hidMouseReportReferenceCreationResult.Error);
        }
        m_hidMouseReportReference = hidMouseReportReferenceCreationResult.Descriptor;

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
                SubscribedHidClientsChanged?.Invoke(null);
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

    private void HidMouseReport_SubscribedClientsChanged(GattLocalCharacteristic sender, object args)
    {
        Debug.WriteLine("Number of clients now registered for mouse notifications: " + sender.SubscribedClients.Count);
        SubscribedHidClientsChanged?.Invoke(sender.SubscribedClients);
    }

    private async Task SendMouseState(bool leftDown, bool rightDown, int mx, int my)
    {
        //lock (m_lock)
        {
            if (!m_initializationFinished)
            {
                return;
            }

            if (m_hidMouseReport.SubscribedClients.Count == 0)
            {
                Debug.WriteLine("No clients are currently subscribed to the mouse report.");
                return;
            }

            var reportValue = new byte[c_sizeOfMouseReportDataInBytes];

            // The first byte of the report data is buttons bitfield.
            reportValue[0] = (byte)((leftDown ? (1 << 0) : 0) | (rightDown ? (1 << 1) : 0));

            reportValue[1] = (byte)(sbyte)mx;
            reportValue[2] = (byte)(sbyte)my;

            Debug.WriteLine("Sending mouse report value notification with data: " + GetStringFromBuffer(reportValue));
            m_lastLeftDown = leftDown;
            m_lastRightDown = rightDown;

            // Waiting for this operation to complete is no longer necessary since now ordering of notifications
            // is guaranteed for each client. Not waiting for it to complete reduces delays and lags.
            // Note that doing this makes us unable to know if the notification failed to be sent.
            await m_hidMouseReport.NotifyValueAsync(reportValue.AsBuffer());
        }
    }
}
