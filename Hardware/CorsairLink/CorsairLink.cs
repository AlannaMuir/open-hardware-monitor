/*
 
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
 
  Copyright (C) 2009-2011 Michael Möller <mmoeller@openhardwaremonitor.org>
	
*/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OpenHardwareMonitor.Hardware.CorsairLink {
    internal class CorsairLinkDevice : Hardware {

        private readonly int portIndex;
        private bool oldClinkDevice; // defines how we can read the registers
        private readonly Sensor[] temperatures;
        private readonly Sensor[] fans;
        private readonly Sensor pump;
        private readonly Control[] controls;
        private readonly List<ISensor> deactivating = new List<ISensor>();
        private readonly HIDDevice hid;
        private readonly int idx;
        private readonly CorsairLink CorsairLink;

        public CorsairLinkDevice(CorsairLink clink, string name, int ident, int index, int channel, HIDDevice hid, ISettings settings)
            : base(name, new Identifier("clink", index.ToString(CultureInfo.InvariantCulture)
            , channel.ToString(CultureInfo.InvariantCulture))
            , settings) {

            this.CorsairLink = clink;
            int tempCount, fanCount, pumpCount;
            switch (ident) {
                case 0x37: oldClinkDevice = true; tempCount = 1; fanCount = 2; pumpCount = 1; break;
                case 0x38: oldClinkDevice = true; tempCount = 4; fanCount = 5; pumpCount = 0; break;
                case 0x3A: oldClinkDevice = true; tempCount = 1; fanCount = 4; pumpCount = 1; break;
                case 0x3B: oldClinkDevice = false; tempCount = 1; fanCount = 4; pumpCount = 1; break;
                case 0x3C: oldClinkDevice = false; tempCount = 1; fanCount = 4; pumpCount = 1; break;
                case 0x3D: oldClinkDevice = false; tempCount = 4; fanCount = 6; pumpCount = 0; break;
                default: throw new Exception("Invalid Device Ident Provided");
            }
            
            temperatures = new Sensor[tempCount];
            fans = new Sensor[fanCount];
            controls = new Control[fanCount];
            if (pumpCount == 1) {
                pump = new Sensor("Pump",
                  0, SensorType.Pump, this, new[] {
                            new ParameterDescription("RPMDivider", "Number of revolutions per minute (RPM) reported versus actual.", 2.0f)
                        }, settings);
                ActivateSensor(pump);
            }

            for (int i = 0; i < tempCount; i++)
                temperatures[i] = new Sensor("Temperature " + (i + 1),
                  i, SensorType.Temperature, this, settings);

            this.hid = hid; this.idx = channel << 4;
            Update();
        }

        void SetFanSpeed(int index, float? value = null) {
            byte mode = 0, pwm = 0; byte[] rpm;
            if (value.HasValue)
            {
                if (value.Value > 100)
                {
                    rpm = BitConverter.GetBytes((ushort)(value.Value));
                    if (oldClinkDevice)
                    {
                        hid.ReadWrite(new byte[] {
                            1, (byte)(8 | idx), (byte)(32 + index * 16), 4, 0,
                            2, (byte)(8 | idx), (byte)(34 + index * 16), rpm[0], rpm[1],
                        });
                    }
                    else
                    {
                        hid.ReadWrite(new byte[] {
                            1, (byte)(6 | idx), 16, (byte)index,
                            2, (byte)(6 | idx), 18, 4,
                            3, (byte)(8 | idx), 20, rpm[0], rpm[1]
                        });
                    }
                }
                else
                {
                    pwm = (byte)(value.Value * 2.55f);
                    mode = 2; // pwm
                    if (oldClinkDevice)
                    {
                        hid.ReadWrite(new byte[] {
                            1, (byte)(8 | idx), (byte)(32 + index * 16), 2, 0,
                            2, (byte)(8 | idx), (byte)(33 + index * 16), pwm, 0,
                        });
                    }
                    else
                    {
                        hid.ReadWrite(new byte[] {
                            1, (byte)(6 | idx), 16, (byte)index,
                            2, (byte)(6 | idx), 18, 2,
                            3, (byte)(6 | idx), 19, pwm
                        });
                    }
                }
            }
            else
            {
                mode = (byte)(oldClinkDevice ? 8 : 6); // default
                hid.ReadWrite(new byte[] {
                    1, (byte)(8 | idx), (byte)(32 + index * 16), (byte)mode, 0
                });
            }
        }

        protected override void ActivateSensor(ISensor sensor) {
            deactivating.Remove(sensor);
            base.ActivateSensor(sensor);
        }

        protected override void DeactivateSensor(ISensor sensor) {
            if (deactivating.Contains(sensor)) {
                deactivating.Remove(sensor);
                base.DeactivateSensor(sensor);
            } else if (active.Contains(sensor)) {
                deactivating.Add(sensor);
            }
        }

        public override HardwareType HardwareType {
            get { return HardwareType.CorsairLink; }
        }

        public override string GetReport() {
            return null;
        }

        public void Open() {
        }

        private void UpdateFan(int i, float rpm, float max, int mode, int pwm, int tgt_rpm) {
            if (fans[i] == null) {
                fans[i] = new Sensor("Fan Channel " + (i + 1), i, SensorType.Fan,
                    this, new[] { new ParameterDescription("MaxRPM", 
                "Maximum revolutions per minute (RPM) of the fan.", 0.0f)
                }, settings);

                controls[i] = new Control(fans[i], settings, 30.0f, 100.0f);
                fans[i].Control = controls[i];
                if ((mode & 14) == 2) {
                    fans[i].Control.SetSoftware(pwm / 2.55f);
                }
                else if ((mode & 14) == 4)
                {
                    fans[i].Control.SetSoftware(tgt_rpm);
                }
                else
                {
                    fans[i].Control.SetDefault();
                }

                controls[i].ControlModeChanged += (cc) =>
                {
                    if (cc.ControlMode == ControlMode.Default)
                        SetFanSpeed(i);
                };
                controls[i].SoftwareControlValueChanged += (cc) =>
                {
                    if (cc.ControlMode == ControlMode.Software)
                        SetFanSpeed(i, cc.SoftwareValue);

                };
            }

            if ((mode & 0x80) > 0) {
                fans[i].Value = rpm;
                fans[i].Parameters[0].Value = max;
                ActivateSensor(fans[i]);
            } else {
                DeactivateSensor(fans[i]);
            }
        }

        public override void Update() {
            if (oldClinkDevice) {
                byte[] data = hid.ReadWrite(new byte[] {
                    1, (byte)(11 | idx), 11, 10, // 4...
                    3, (byte)(11 | idx), 16, 10, // 17...
                    5, (byte)(9 | idx), 7, // 1d
                    6, (byte)(9 | idx), 8, // 21
                    7, (byte)(9 | idx), 9, // 25
                    8, (byte)(9 | idx), 10, // 29
                });
                if (data == null) return;
                for (int i = 0; i < temperatures.Length; i++) {
                    float temp = BitConverter.ToInt16(data, i * 4 + 29) / 256.0f;
                    if (temp > 0.0f) {
                        temperatures[i].Value = temp;
                        ActivateSensor(temperatures[i]);
                    } else {
                        DeactivateSensor(temperatures[i]);
                    }
                }

                byte[] mpwm = hid.ReadWrite(new byte[] {
                    9, (byte)(11 | idx), 32, 6, // 4
                    10, (byte)(11 | idx), 48, 6, // 13
                    11, (byte)(11 | idx), 64, 6, // 22
                    12, (byte)(11 | idx), 80, 6, // 31
                    13, (byte)(11 | idx), 96, 6, // 40
                });
                if (mpwm == null) return;
                for (int i = 0; i < fans.Length; i++) {
                    float rpm = (float)BitConverter.ToUInt16(data, 4 + i * 2);
                    float max = (float)BitConverter.ToUInt16(data, 17 + i * 2);
                    int mode = mpwm[4 + 9 * i];
                    int pwm = mpwm[6 + 9 * i];
                    int tgt = BitConverter.ToUInt16(mpwm, 8 + 9 * i);
                    UpdateFan(i, rpm, max, mode, pwm, tgt);
                }
                    
            } else {
                for (int i = 0; i < fans.Length; i++) {
                    byte[] data = hid.ReadWrite(new byte[] {
                        1,(byte)(6 | idx), 16, (byte)i, // write fan channel
                        2,(byte)(7 | idx), 18, // 5 fan mode
                        3,(byte)(7 | idx), 19, // 8 fan pwm
                        4,(byte)(9 | idx), 22, // b/c rpm
                        5,(byte)(9 | idx), 23, // f/10 max
                        6,(byte)(7 | idx), 20, // 13/14 fan target rpm

                    });
                    if (data == null) return;
                    int mode = data[5], pwm = data[8];
                    float rpm = (float)BitConverter.ToUInt16(data, 11);
                    float max = (float)BitConverter.ToUInt16(data, 15);
                    ushort tgt = BitConverter.ToUInt16(data, 19);
                    UpdateFan(i, rpm, max, mode, pwm, tgt);
                }

                for (int i = 0; i < temperatures.Length; i++) {
                    byte[] data = hid.ReadWrite(new byte[] {
                        1,(byte)(6 | idx), 12, (byte)i, // write temp channel
                        2,(byte)(7 | idx), 14, // 5 temp value
                    });
                    if (data == null) return;
                    float temp = (float)BitConverter.ToInt16(data, 5) / 256.0f;
                    if (temp > 0.0f) {
                        temperatures[i].Value = temp;
                        ActivateSensor(temperatures[i]);
                    } else {
                        DeactivateSensor(temperatures[i]);
                    }

                }

            }
        }

        public override IHardware Parent {
            get { return CorsairLink; }
        }

        public override void Close() {
            for(int i=0; i<temperatures.Length; i++) 
                DeactivateSensor(temperatures[i]);

            for(int i=0; i<fans.Length; i++) 
                DeactivateSensor(fans[i]);

            //for(int i=0; i<controls.Length; i++) 
                //DeactivateSensor(controls[i]);
            
            DeactivateSensor(pump);
            base.Close();
        }

    }
}
