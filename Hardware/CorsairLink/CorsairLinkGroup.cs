/*
 
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
 
  Copyright (C) 2009-2012 Michael Möller <mmoeller@openhardwaremonitor.org>
	
*/

using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;

namespace OpenHardwareMonitor.Hardware.CorsairLink {
    internal class CorsairLinkGroup : IGroup {
        internal static readonly List<CorsairLink> hardware =
            new List<CorsairLink>();

        private ISettings settings;

        private Thread CorsairLinkGroup_Thread = new Thread((me) =>
        {
            // The top level contains all USB devices such as the
            // USB Commander and H80i's
            var _this = me as CorsairLinkGroup;
            while (true)
            {
                var paths = HIDDevice.FindDevices(0x1B1C, 0x0C04);
                foreach(var path in paths) {
                    bool is_new_device = false;
                    var device = HIDDevice.Open(path, out is_new_device);
                    if (is_new_device)
                    {
                        try
                        {
                            int index = hardware.Count;
                            var hw = new CorsairLink(device, index, _this.settings);
                            hardware.Add(hw);
                        }
                        catch { }
                    }
                    Thread.Sleep(100);
                }
                Thread.Sleep(1000);
            }
        });

        public CorsairLinkGroup(ISettings settings) {
            this.settings = settings;
            CorsairLinkGroup_Thread.IsBackground = true;
            CorsairLinkGroup_Thread.Start(this);
        }

        public string GetReport() {
            return null;
        }

        public IHardware[] Hardware {
            get {
                return hardware.ToArray();
            }
        }

        public void Close()
        {
            foreach (CorsairLink clink in hardware)
                clink.Close();
        }
    }

    internal class CorsairLink : IHardware {
        private readonly string name = "Corsair Link";
        private string customName;
        private readonly List<CorsairLinkDevice> hardware =
           new List<CorsairLinkDevice>();
        private readonly ISettings settings;

#pragma warning disable 67
        public event SensorEventHandler SensorAdded;
        public event SensorEventHandler SensorRemoved;
#pragma warning restore 67

        //public HIDDevice Hid;

        private void AddDevice(HIDDevice hid, int count, int channel, ISettings settings) {
            int idx = channel << 4;
            string name = "";
            // data[2] is ident, data[6] is major/minor data[5] revision
            byte[] data = hid.ReadWrite(new byte[] { 1, (byte)(7 | idx), 0, 2, (byte)(9 | idx), 1 });
            if (data == null) return;

            int ident = data[3];
            int major = data[7] >> 4, minor = data[7] & 0xF, rev = data[6];
            string version = major + "." + minor + "." + rev;
            switch (ident) {
                case 0x37: name = "Hydro H80"; break;
                case 0x38: name = "Cooling Node"; break;
                case 0x3A: name = "Hydro H100"; break;
                case 0x3B: name = "Hydro H80i"; break;
                case 0x3C: name = "Hydro H100i"; break;
                case 0x3D: name = "Commander Mini"; break;
                default: return; // other node types aren't supported (yet?)
            }
            name += " v" + version;
            hardware.Add(new CorsairLinkDevice(this, name, ident, count, channel, hid, settings));
        }

        public CorsairLink(HIDDevice hid, Int32 index, ISettings settings)
        {
            this.customName = settings.GetValue(
            new Identifier(Identifier, "name").ToString(), name);
            this.settings = settings;

            byte[] data = hid.ReadWrite(new byte[] { 1, 79 });
            if (data == null) throw new Exception("Cannot create device");

            AddDevice(hid, index, 0, settings);
            for (int i = 1; i < 8; i++)
            {
                if (data[3 + i] != 5) continue;
                AddDevice(hid, index, i, settings);
            }
        }

        public CorsairLink()
        {
            //CorsairLinkGroup.ClinkHardware;
            // TODO: Complete member initialization
        }
        public virtual IHardware Parent {
            get { return null; }
        }

        public IHardware[] Hardware {
            get {
                return hardware.ToArray();
            }
        }
        public string GetReport() {
            // return null if there's no report
            // if (report.Length == 0) return null
            return null;
        }

        public void Close() {
            foreach (var h in hardware)
            {
                h.Close();
            }
            //Hid.Close();
        }

        public string Name {
            get {
                return customName;
            }
            set {
                if (!string.IsNullOrEmpty(value))
                    customName = value;
                else
                    customName = name;
                settings.SetValue(new Identifier(Identifier, "name").ToString(),
                  customName);
            }
        }

        public Identifier Identifier {
            get { return new Identifier("clink"); }
        }

        public IHardware[] SubHardware {
            get { return hardware.ToArray(); }
        }

        public ISensor[] Sensors {
            get { return new ISensor[0]; }
        }

        public void Accept(IVisitor visitor) {
            if (visitor == null) // 1
                throw new ArgumentNullException("visitor");
            visitor.VisitHardware(this);
        }

        public void Traverse(IVisitor visitor) {
            foreach (IHardware h in hardware)
                h.Accept(visitor);
        }

        public HardwareType HardwareType {
            get { return HardwareType.CorsairLink; }
        }

        public void Update() {
            
            Console.WriteLine("..."); // 2
        }
    }
}
