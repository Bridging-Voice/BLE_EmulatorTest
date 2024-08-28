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
            0x15, 0x81,        //     Logical Minimum (-127)
            0x25, 0x7F,        //     Logical Maximum (127)
            0x75, 0x08,        //     Report Size (8)
            0x95, 0x03,        //     Report Count (3)
            0x81, 0x06,        //     Input (Data,Var,Rel,No Wrap,Linear,Preferred State,No Null Position)
            0xC0,              //   End Collection
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

    private static readonly GattLocalCharacteristicParameters c_batteryLevelParameters = new GattLocalCharacteristicParameters
    {
        CharacteristicProperties = GattCharacteristicProperties.Read | GattCharacteristicProperties.Notify,
        ReadProtectionLevel = GattProtectionLevel.Plain
    };

    private static readonly uint c_sizeOfMouseReportDataInBytes = 0x4;

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

    public async Task Move(int mx, int my, int wheel)
    {
        try
        {
            await SendMouseState(m_lastLeftDown, m_lastRightDown, mx, my, wheel);
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

    public async Task Click()
    {
        try
        {
            await SendMouseState(true, false, 0, 0, 0);
            await Task.Delay(40);
            await SendMouseState(false, false, 0, 0, 0);
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

    private async Task SendMouseState(bool leftDown, bool rightDown, int mx, int my, int wheel)
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
            reportValue[3] = (byte)(sbyte)wheel;

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
