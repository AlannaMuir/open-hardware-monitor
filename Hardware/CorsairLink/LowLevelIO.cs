using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OpenHardwareMonitor.Hardware.CorsairLink {

    /// <summary>
    /// Win32Usb provides all low-level access to the Win32 API within
    /// our .Net applications -- these are mostly struct defines and
    /// dllimports to bind to kernel.dll, hid,dll and setupapi.dll
    /// </summary>
    internal class Win32
    {
        #region Structures
        /// <summary>
        /// An overlapped structure used for overlapped IO operations. The structure is
        /// only used by the OS to keep state on pending operations. You don't need to fill anything in if you
        /// unless you want a Windows event to fire when the operation is complete.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct Overlapped
        {
            public uint Internal;
            public uint InternalHigh;
            public uint Offset;
            public uint OffsetHigh;
            public IntPtr Event;
        }
        /// <summary>
        /// Provides details about a single USB device
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct DeviceInterfaceData
        {
            public int Size;
            public Guid InterfaceClassGuid;
            public int Flags;
            public int Reserved;
        }
        /// <summary>
        /// Provides the capabilities of a HID device
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct HidCaps
        {
            public short Usage;
            public short UsagePage;
            public short InputReportByteLength;
            public short OutputReportByteLength;
            public short FeatureReportByteLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
            public short[] Reserved;
            public short NumberLinkCollectionNodes;
            public short NumberInputButtonCaps;
            public short NumberInputValueCaps;
            public short NumberInputDataIndices;
            public short NumberOutputButtonCaps;
            public short NumberOutputValueCaps;
            public short NumberOutputDataIndices;
            public short NumberFeatureButtonCaps;
            public short NumberFeatureValueCaps;
            public short NumberFeatureDataIndices;
        }
        /// <summary>
        /// Access to the path for a device
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct DeviceInterfaceDetailData
        {
            public int Size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string DevicePath;
        }
        /// <summary>
        /// Used when registering a window to receive messages about devices added or removed from the system.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 1)]
        public class DeviceBroadcastInterface
        {
            public int Size;
            public int DeviceType;
            public int Reserved;
            public Guid ClassGuid;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string Name;
        }
        #endregion

        #region Constants
        /// <summary>Windows message sent when a device is inserted or removed</summary>
        public const int WM_DEVICECHANGE = 0x0219;
        /// <summary>WParam for above : A device was inserted</summary>
        public const int DEVICE_ARRIVAL = 0x8000;
        /// <summary>WParam for above : A device was removed</summary>
        public const int DEVICE_REMOVECOMPLETE = 0x8004;
        /// <summary>Used in SetupDiClassDevs to get devices present in the system</summary>
        protected const int DIGCF_PRESENT = 0x02;
        /// <summary>Used in SetupDiClassDevs to get device interface details</summary>
        protected const int DIGCF_DEVICEINTERFACE = 0x10;
        /// <summary>Used when registering for device insert/remove messages : specifies the type of device</summary>
        protected const int DEVTYP_DEVICEINTERFACE = 0x05;
        /// <summary>Used when registering for device insert/remove messages : we're giving the API call a window handle</summary>
        protected const int DEVICE_NOTIFY_WINDOW_HANDLE = 0;
        /// <summary>Purges Win32 transmit buffer by aborting the current transmission.</summary>
        protected const uint PURGE_TXABORT = 0x01;
        /// <summary>Purges Win32 receive buffer by aborting the current receive.</summary>
        protected const uint PURGE_RXABORT = 0x02;
        /// <summary>Purges Win32 transmit buffer by clearing it.</summary>
        protected const uint PURGE_TXCLEAR = 0x04;
        /// <summary>Purges Win32 receive buffer by clearing it.</summary>
        protected const uint PURGE_RXCLEAR = 0x08;
        /// <summary>CreateFile : Open file for read</summary>
        protected const uint GENERIC_READ = 0x80000000;
        /// <summary>CreateFile : Open file for write</summary>
        protected const uint GENERIC_WRITE = 0x40000000;

        /// <summary>CreateFile : subsequent open operations on a file or device can read</summary>
        protected const uint FILE_SHARE_READ = 0x00000001;
        /// <summary>CreateFile : subsequent open operations on a file or device can write</summary>
        protected const uint FILE_SHARE_WRITE = 0x00000002;
        /// <summary>CreateFile : subsequent open operations on a file or device can delete</summary>
        protected const uint FILE_SHARE_DELETE = 0x00000004;

        /// <summary>CreateFile : Open handle for overlapped operations</summary>
        protected const uint FILE_FLAG_OVERLAPPED = 0x40000000;
        /// <summary>CreateFile : Resource to be "created" must exist</summary>
        protected const uint OPEN_EXISTING = 3;
        /// <summary>ReadFile/WriteFile : Overlapped operation is incomplete.</summary>
        protected const uint ERROR_IO_PENDING = 997;
        /// <summary>Infinite timeout</summary>
        protected const uint INFINITE = 0xFFFFFFFF;
        /// <summary>Simple representation of a null handle : a closed stream will get this handle. Note it is public for comparison by higher level classes.</summary>
        public static IntPtr NullHandle = IntPtr.Zero;
        /// <summary>Simple representation of the handle returned when CreateFile fails.</summary>
        protected static IntPtr InvalidHandleValue = new IntPtr(-1);
        #endregion

        #region P/Invoke
        /// <summary>
        /// Check that two buffers are identical.
        /// </summary>
        /// <param name="b1">First byte array.</param>
        /// <param name="b2">Second byte array.</param>
        /// <param name="count">The number of bytes to check.</param>
        /// <returns>Zero if equal, any value if not.</returns>
        [DllImport("msvcrt.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int memcmp(byte[] b1, byte[] b2, long count);

        /// <summary>
        /// deletes all pending input reports in a top-level collection's input queue
        /// </summary>
        /// <param name="gHid">An in parameter to take the Guid</param>
        [DllImport("hid.dll", SetLastError = true)]
        protected static extern bool HidD_FlushQueue(Guid gHid);
        /// <summary>
        /// Gets the GUID that Windows uses to represent HID class devices
        /// </summary>
        /// <param name="gHid">An out parameter to take the Guid</param>
        [DllImport("hid.dll", SetLastError = true)]
        protected static extern void HidD_GetHidGuid(out Guid gHid);
        /// <summary>
        /// Allocates an InfoSet memory block within Windows that contains details of devices.
        /// </summary>
        /// <param name="gClass">Class guid (e.g. HID guid)</param>
        /// <param name="strEnumerator">Not used</param>
        /// <param name="hParent">Not used</param>
        /// <param name="nFlags">Type of device details required (DIGCF_ constants)</param>
        /// <returns>A reference to the InfoSet</returns>
        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, [MarshalAs(UnmanagedType.LPStr)] string strEnumerator, IntPtr hParent, uint nFlags);
        /// <summary>
        /// Frees InfoSet allocated in call to above.
        /// </summary>
        /// <param name="lpInfoSet">Reference to InfoSet</param>
        /// <returns>true if successful</returns>
        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern int SetupDiDestroyDeviceInfoList(IntPtr lpInfoSet);
        /// <summary>
        /// Gets the DeviceInterfaceData for a device from an InfoSet.
        /// </summary>
        /// <param name="lpDeviceInfoSet">InfoSet to access</param>
        /// <param name="nDeviceInfoData">Not used</param>
        /// <param name="gClass">Device class guid</param>
        /// <param name="nIndex">Index into InfoSet for device</param>
        /// <param name="oInterfaceData">DeviceInterfaceData to fill with data</param>
        /// <returns>True if successful, false if not (e.g. when index is passed end of InfoSet)</returns>
        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiEnumDeviceInterfaces(IntPtr lpDeviceInfoSet, uint nDeviceInfoData, ref Guid gClass, uint nIndex, ref DeviceInterfaceData oInterfaceData);
        /// <summary>
        /// SetupDiGetDeviceInterfaceDetail - two of these, overloaded because they are used together in slightly different
        /// ways and the parameters have different meanings.
        /// Gets the interface detail from a DeviceInterfaceData. This is pretty much the device path.
        /// You call this twice, once to get the size of the struct you need to send (nDeviceInterfaceDetailDataSize=0)
        /// and once again when you've allocated the required space.
        /// </summary>
        /// <param name="lpDeviceInfoSet">InfoSet to access</param>
        /// <param name="oInterfaceData">DeviceInterfaceData to use</param>
        /// <param name="lpDeviceInterfaceDetailData">DeviceInterfaceDetailData to fill with data</param>
        /// <param name="nDeviceInterfaceDetailDataSize">The size of the above</param>
        /// <param name="nRequiredSize">The required size of the above when above is set as zero</param>
        /// <param name="lpDeviceInfoData">Not used</param>
        /// <returns></returns>
        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr lpDeviceInfoSet, ref DeviceInterfaceData oInterfaceData, IntPtr lpDeviceInterfaceDetailData, uint nDeviceInterfaceDetailDataSize, ref uint nRequiredSize, IntPtr lpDeviceInfoData);
        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr lpDeviceInfoSet, ref DeviceInterfaceData oInterfaceData, ref DeviceInterfaceDetailData oDetailData, uint nDeviceInterfaceDetailDataSize, ref uint nRequiredSize, IntPtr lpDeviceInfoData);
        /// <summary>
        /// Registers a window for device insert/remove messages
        /// </summary>
        /// <param name="hwnd">Handle to the window that will receive the messages</param>
        /// <param name="lpInterface">DeviceBroadcastInterrface structure</param>
        /// <param name="nFlags">set to DEVICE_NOTIFY_WINDOW_HANDLE</param>
        /// <returns>A handle used when unregistering</returns>
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr RegisterDeviceNotification(IntPtr hwnd, DeviceBroadcastInterface oInterface, uint nFlags);
        /// <summary>
        /// Unregister from above.
        /// </summary>
        /// <param name="hHandle">Handle returned in call to RegisterDeviceNotification</param>
        /// <returns>True if success</returns>
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern bool UnregisterDeviceNotification(IntPtr hHandle);
        /// <summary>
        /// Gets details from an open device. Reserves a block of memory which must be freed.
        /// </summary>
        /// <param name="hFile">Device file handle</param>
        /// <param name="lpData">Reference to the preparsed data block</param>
        /// <returns></returns>
        [DllImport("hid.dll", SetLastError = true)]
        protected static extern bool HidD_GetPreparsedData(IntPtr hFile, out IntPtr lpData);
        /// <summary>
        /// Frees the memory block reserved above.
        /// </summary>
        /// <param name="pData">Reference to preparsed data returned in call to GetPreparsedData</param>
        /// <returns></returns>
        [DllImport("hid.dll", SetLastError = true)]
        protected static extern bool HidD_FreePreparsedData(ref IntPtr pData);
        /// <summary>
        /// Gets a device's capabilities from the preparsed data.
        /// </summary>
        /// <param name="lpData">Preparsed data reference</param>
        /// <param name="oCaps">HidCaps structure to receive the capabilities</param>
        /// <returns>True if successful</returns>
        [DllImport("hid.dll", SetLastError = true)]
        protected static extern int HidP_GetCaps(IntPtr lpData, out HidCaps oCaps);
        /// <summary>
        /// Creates/opens a file, serial port, USB device... etc
        /// </summary>
        /// <param name="strName">Path to object to open</param>
        /// <param name="nAccess">Access mode. e.g. Read, write</param>
        /// <param name="nShareMode">Sharing mode</param>
        /// <param name="lpSecurity">Security details (can be null)</param>
        /// <param name="nCreationFlags">Specifies if the file is created or opened</param>
        /// <param name="nAttributes">Any extra attributes? e.g. open overlapped</param>
        /// <param name="lpTemplate">Not used</param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        protected static extern IntPtr CreateFile([MarshalAs(UnmanagedType.LPStr)] string strName, uint nAccess, uint nShareMode, IntPtr lpSecurity, uint nCreationFlags, uint nAttributes, IntPtr lpTemplate);
        /// <summary>
        /// Closes a window handle. File handles, event handles, mutex handles... etc
        /// </summary>
        /// <param name="hFile">Handle to close</param>
        /// <returns>True if successful.</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        protected static extern int CloseHandle(IntPtr hFile);
        #endregion

        #region Public methods
        /// <summary>
        /// Registers a window to receive windows messages when a device is inserted/removed. Need to call this
        /// from a form when its handle has been created, not in the form constructor. Use form's OnHandleCreated override.
        /// </summary>
        /// <param name="hWnd">Handle to window that will receive messages</param>
        /// <param name="gClass">Class of devices to get messages for</param>
        /// <returns>A handle used when unregistering</returns>
        public static IntPtr RegisterForUsbEvents(IntPtr hWnd, Guid gClass)
        {
            DeviceBroadcastInterface oInterfaceIn = new DeviceBroadcastInterface();
            oInterfaceIn.Size = Marshal.SizeOf(oInterfaceIn);
            oInterfaceIn.ClassGuid = gClass;
            oInterfaceIn.DeviceType = DEVTYP_DEVICEINTERFACE;
            oInterfaceIn.Reserved = 0;
            return RegisterDeviceNotification(hWnd, oInterfaceIn, DEVICE_NOTIFY_WINDOW_HANDLE);
        }

        public static bool ArraysEqual(byte[] array1, byte[] array2)
        {
            if (array1 == null || array2 == null) return false;
            if (array1.Length != array2.Length) return false;
            return memcmp(array1, array2, array1.Length) == 0;
        }

        /// <summary>
        /// Unregisters notifications. Can be used in form dispose
        /// </summary>
        /// <param name="hHandle">Handle returned from RegisterForUSBEvents</param>
        /// <returns>True if successful</returns>
        public static bool UnregisterForUsbEvents(IntPtr hHandle)
        {
            return UnregisterDeviceNotification(hHandle);
        }
        /// <summary>
        /// Helper to get the HID guid.
        /// </summary>
        public static Guid HIDGuid
        {
            get
            {
                Guid gHid;
                HidD_GetHidGuid(out gHid);
                return gHid;
            }
        }
        #endregion
    }

    internal class HIDDevice : Win32
    {
        private IntPtr UsbHandle;
        private Int16 ReportLength;
        private FileStream UsbStream;
        private AutoResetEvent UsbComplete;
        public String UsbPath;

        public Boolean IsConnected
        {
            get
            {
                return Enumerate().Count(p => p == UsbPath) > 0;
            }
        }

        /// <summary>
        /// Initialises and opens the device
        /// </summary>
        /// <param name="strPath">Path to the device</param>
        private HIDDevice(string strPath)
        {
            UsbPath = strPath;
            // Create the file from the device path
            UsbHandle = CreateFile(strPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, IntPtr.Zero);
            if (UsbHandle != InvalidHandleValue)	// if the open worked...
            {
                IntPtr lpData;
                if (HidD_GetPreparsedData(UsbHandle, out lpData))	// get windows to read the device data into an internal buffer
                {
                    try
                    {
                        HidCaps oCaps;
                        HidP_GetCaps(lpData, out oCaps);	// extract the device capabilities from the internal buffer
                        //InputReportLength = oCaps.InputReportByteLength;	// get the input...
                        ReportLength = oCaps.OutputReportByteLength;	// ... and output report lengths
                        UsbStream = new FileStream(UsbHandle, FileAccess.Read | FileAccess.Write, true, ReportLength, true);	// wrap the file handle in a .Net file stream

                        //BeginAsyncRead();	// kick off the first asynchronous read
                    }
                    catch
                    {
                        throw; // raise unhandled exceptions to the caller
                    }
                    finally
                    {
                        HidD_FreePreparsedData(ref lpData);	// before we quit the funtion, we must free the internal buffer reserved in GetPreparsedData
                    }
                }
                else	// GetPreparsedData failed? Chuck an exception
                {
                    throw new Exception("GetPreparsedData failed");
                }
            }
            else	// File open failed? Chuck an exception
            {
                UsbHandle = IntPtr.Zero;
                throw new Exception("Failed to create device file");
            }
            UsbComplete = new AutoResetEvent(false);
        }

        /// <summary>
        /// Helper method to return the device path given a DeviceInterfaceData structure and an InfoSet handle.
        /// Used in 'FindDevice' so check that method out to see how to get an InfoSet handle and a DeviceInterfaceData.
        /// </summary>
        /// <param name="hInfoSet">Handle to the InfoSet</param>
        /// <param name="oInterface">DeviceInterfaceData structure</param>
        /// <returns>The device path or null if there was some problem</returns>
        private static string GetDevicePath(IntPtr hInfoSet, ref DeviceInterfaceData oInterface)
        {
            uint nRequiredSize = 0;
            // Get the device interface details
            if (!SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, IntPtr.Zero, 0, ref nRequiredSize, IntPtr.Zero))
            {
                DeviceInterfaceDetailData oDetail = new DeviceInterfaceDetailData();
                // hardcoded to 5! Sorry, but this works and trying more future proof versions by setting the size to the struct sizeof failed miserably. If you manage to sort it, mail me! Thx
                oDetail.Size = 5;
                if (SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, ref oDetail, nRequiredSize, ref nRequiredSize, IntPtr.Zero))
                {
                    return oDetail.DevicePath;
                }
            }
            return null;
        }

        private static List<String> Enumerate()
        {
            List<String> paths = new List<String>();
            Guid gHid;
            HidD_GetHidGuid(out gHid);	// next, get the GUID from Windows that it uses to represent the HID USB interface
            IntPtr hInfoSet = SetupDiGetClassDevs(ref gHid, null, IntPtr.Zero, DIGCF_DEVICEINTERFACE | DIGCF_PRESENT);	// this gets a list of all HID devices currently connected to the computer (InfoSet)
            try
            {
                DeviceInterfaceData oInterface = new DeviceInterfaceData();	// build up a device interface data block
                oInterface.Size = Marshal.SizeOf(oInterface);
                // Now iterate through the InfoSet memory block assigned within Windows in the call to SetupDiGetClassDevs
                // to get device details for each device connected
                int nIndex = 0;
                while (SetupDiEnumDeviceInterfaces(hInfoSet, 0, ref gHid, (uint)nIndex, ref oInterface))	// this gets the device interface information for a device at index 'nIndex' in the memory block
                {
                    string strDevicePath = GetDevicePath(hInfoSet, ref oInterface);	// get the device path (see helper method 'GetDevicePath')
                    paths.Add(strDevicePath);
                    nIndex++;	// if we get here, we didn't find our device. So move on to the next one.
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                // Before we go, we have to free up the InfoSet memory reserved by SetupDiGetClassDevs
                SetupDiDestroyDeviceInfoList(hInfoSet);
            }
            return paths;
        }

        private static Dictionary<String, HIDDevice> devices =
            new Dictionary<String, HIDDevice>();

        /// <summary>
        /// Finds a device given its PID and VID, and opens it
        /// </summary>
        /// <param name="nVid">Vendor id for device (VID)</param>
        /// <param name="nPid">Product id for device (PID)</param>
        /// <param name="oType">Type of device class to create</param>
        /// <returns>A new device class of the given type or null</returns>
        public static List<String> FindDevices(int nVid, int nPid)
        {
            //var devices = new List<HIDDevice>();
            string strSearch = string.Format("vid_{0:x4}&pid_{1:x4}", nVid, nPid);
            var paths = Enumerate().Where(p => p.Contains(strSearch));
            foreach (var path in paths)
            {
                if (!devices.ContainsKey(path))
                    devices.Add(path, null);
            }
            
            return paths.ToList();
            //return null;	// oops, didn't find our device
        }

        public static HIDDevice Open(String path, out Boolean newDevice)
        {
            bool _new = false;
            if (!devices.ContainsKey(path) || devices[path] == null)
            {
                devices[path] = new HIDDevice(path);
                _new = true;
            }
            newDevice = _new;
            return devices[path];
        }

        private void CompleteAsyncRead(IAsyncResult ar)
        {
            if (UsbStream != null)
            {
                try
                {
                    UsbStream.EndRead(ar);
                }
                catch (ObjectDisposedException)
                {
                    Close();
                }
                catch (IOException)
                {
                    UsbStream.Flush();
                }
                finally
                {
                    UsbComplete.Set();
                }
            }
            else
            {
                Close();
            }
        }

        public void Close()
        {
            if (UsbStream != null)
            {
                UsbStream.Close();
                UsbStream = null;
            }
            devices[UsbPath] = null;
        }

        /// <summary>
        /// Writes one packet to the device expecting a (near) immediate response.
        /// </summary>
        /// <param name="out_packet">Packet to write to the device.</param>
        /// <returns>The packet read from the device, or null on any error.</returns>
        public byte[] ReadWrite(byte[] out_packet)
        {
            var UsbBuffer = new byte[ReportLength];
            UsbBuffer[0] = 0; UsbBuffer[1] = (byte)out_packet.Length;
            int UsbLength = Math.Min(out_packet.Length, ReportLength - 1);
            Buffer.BlockCopy(out_packet, 0, UsbBuffer, 2, UsbLength);

            if (UsbStream == null) return null;

            try
            {
                UsbComplete.Reset();
                UsbStream.BeginRead(UsbBuffer, 0, ReportLength, CompleteAsyncRead, null);
                UsbStream.Write(UsbBuffer, 0, ReportLength);
                if (!UsbComplete.WaitOne(500)) throw new IOException("Timout waiting for device.");
                return UsbBuffer;
            }
            catch (ObjectDisposedException)
            {
                 Close();
            }
            catch (IOException)
            {
                UsbStream.Flush();
            }
            catch { throw; }
            return null;
        }
    }
}
