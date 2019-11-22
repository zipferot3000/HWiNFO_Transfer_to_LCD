using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.MemoryMappedFiles;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;


namespace HWiNFO_Transfer_to_LCD
{
    public partial class Service1 : ServiceBase
    {
        const string HWiNFO_SHARED_MEM_FILE_NAME = "Global\\HWiNFO_SENS_SM2";
        const int HWiNFO_SENSORS_STRING_LEN = 128;
        const int HWiNFO_UNIT_STRING_LEN = 16;
        static MemoryMappedFile mmf;
        static EventLog eventLog;
        static MemoryMappedViewAccessor accessor;
        static HWiNFO_SHARED_MEM HWiNFOMemory;
        static List<HWiNFO_ELEMENT> data_arr;
        static List<string> values_to_send;
        static SerialPort serialPort;
        private Timer timer;
        static bool port_opened = false;
        static Properties.Settings settings = Properties.Settings.Default;

        public Service1()
        {
            InitializeComponent();
            eventLog = new EventLog();
            try 
            {
                EventLog.CreateEventSource(settings.SysLogSourceName, settings.SysLogName);
            }
            catch
            {

            }
            eventLog.Source = settings.SysLogSourceName;
            eventLog.Log = settings.SysLogName;
            timer = new Timer();
            timer.Elapsed += new ElapsedEventHandler(timer_Tick);
            timer.Interval = 1000;
        }

        protected override void OnStart(string[] args)
        {
            openSerialPort();
            timer.Start();
        }

        protected override void OnStop()
        {
            this.timer.Stop();
            if (serialPort.IsOpen)
            {
                try { serialPort.Close(); } catch { }
            }
        }

        private static bool HWiNFOisStarted()
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
                if (Regex.IsMatch(process.ProcessName, settings.HWiNFO_PROCCESS_NAME))
                {
                    return true;
                }
            }
            return false;
        }

        private static void readMemory()
        {
            try
            {
                mmf = MemoryMappedFile.OpenExisting(HWiNFO_SHARED_MEM_FILE_NAME, MemoryMappedFileRights.Read);
                accessor = mmf.CreateViewAccessor(0L, (long)Marshal.SizeOf(typeof(HWiNFO_SHARED_MEM)), MemoryMappedFileAccess.Read);
                accessor.Read<HWiNFO_SHARED_MEM>(0L, out HWiNFOMemory);
                data_arr = new List<HWiNFO_ELEMENT>();
                for (uint index = 0; index < HWiNFOMemory.dwNumReadingElements; ++index)
                {
                    using (MemoryMappedViewStream viewStream = mmf.CreateViewStream((long)(HWiNFOMemory.dwOffsetOfReadingSection + index * HWiNFOMemory.dwSizeOfReadingElement), (long)HWiNFOMemory.dwSizeOfReadingElement, MemoryMappedFileAccess.Read))
                    {
                        byte[] buffer = new byte[new IntPtr(HWiNFOMemory.dwSizeOfReadingElement).ToInt32()];
                        viewStream.Read(buffer, 0, (int)HWiNFOMemory.dwSizeOfReadingElement);
                        GCHandle gcHandle = GCHandle.Alloc((object)buffer, GCHandleType.Pinned);
                        HWiNFO_ELEMENT structure = (HWiNFO_ELEMENT)Marshal.PtrToStructure(gcHandle.AddrOfPinnedObject(), typeof(HWiNFO_ELEMENT));
                        gcHandle.Free();
                        data_arr.Add(structure);
                    }
                }
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry("An error occured while opening the HWiNFO shared memory: " + ex.Message, EventLogEntryType.Error);
            }
        }

        private static string makeSerialData()
        {
            try
            {
                string resolve = "", tmp_str = "";
                int tmp = 0;
                double tmp2 = 0, tmp3 = 0;
                values_to_send = new List<string>();

                //set time
                values_to_send.Add(DateTime.Now.ToString("HH:mm"));

                //set date
                values_to_send.Add(DateTime.Now.ToString("dd.MM.yy"));

                //CPU Package temp
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.CPUPackageTempID && d.dwSensorIndex == settings.CPUPackageTempIndex)).Value)).ToString());

                //Total CPU Usage
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.TotalCPUUsageID && d.dwSensorIndex == settings.TotalCPUUsageIndex)).Value);
                values_to_send.Add(getStr(tmp, "%"));

                //CPU Package Power
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.CPUPackagePowerID && d.dwSensorIndex == settings.CPUPackagePowerIndex)).Value);
                values_to_send.Add(getStr(tmp, "W"));

                //Fan3
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.CPUFanID && d.dwSensorIndex == settings.CPUFanIndex)).Value);
                values_to_send.Add(getRPM(tmp));




                //GPU Temp
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.GPUTempID && d.dwSensorIndex == settings.GPUIndex)).Value)).ToString());

                //GPU Core Load
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.GPUCoreLoadID && d.dwSensorIndex == settings.GPUIndex)).Value);
                tmp_str = getStr(tmp, "%");
                values_to_send.Add(tmp_str == "100%" ? "100" : tmp_str.Substring(0,3));

                //GPU Memory Usage
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.GPUMemoryUsageID && d.dwSensorIndex == settings.GPUIndex)).Value);
                tmp_str = getStr(tmp, "%");
                values_to_send.Add(tmp_str == "100%" ? "100" : tmp_str.Substring(0, 3));

                //GPU Power
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.GPUPowerID && d.dwSensorIndex == settings.GPUIndex)).Value);
                values_to_send.Add(getStr(tmp, "W"));

                //GPU Fan1
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.GPUFanID && d.dwSensorIndex == settings.GPUIndex)).Value);
                values_to_send.Add(getRPM(tmp));




                //Memory Usage
                tmp = (int)((data_arr.Find(d => d.dwReadingID == settings.MemoryUsageID && d.dwSensorIndex == settings.MemoryUsageIndex)).Value);
                values_to_send.Add(getStr(tmp, "%"));

                //DIMM[0] Temperature
                int temp = (int)((data_arr.Find(d => d.dwReadingID == settings.DIMM_0_TemperatureID && d.dwSensorIndex == settings.DIMM_0_TemperatureIndex)).Value);
                int temp2 = (int)((data_arr.Find(d => d.dwReadingID == settings.DIMM_2_TemperatureID && d.dwSensorIndex == settings.DIMM_2_TemperatureIndex)).Value);

                values_to_send.Add(Math.Max(temp, temp2).ToString());



                //PSU Power
                double pef = ((data_arr.Find(d => d.dwReadingID == settings.PSUPowerEfficiencyID && d.dwSensorIndex == settings.PSUPowerEfficiencyIndex)).Value / 100);
                double pow = (data_arr.Find(d => d.dwReadingID == settings.PSUPowerID && d.dwSensorIndex == settings.PSUPowerIndex)).Value;
                tmp = (int)(pow / pef);
                values_to_send.Add(getStr(tmp, "W"));
                
                //PSU Temperature
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.PSUTemperatureID && d.dwSensorIndex == settings.PSUTemperatureIndex)).Value)).ToString());



                //Disk temps
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.DiskTempID && d.dwSensorIndex == settings.Disk1TempIndex)).Value)).ToString());
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.DiskTempID && d.dwSensorIndex == settings.Disk2TempIndex)).Value)).ToString());
                values_to_send.Add(((int)((data_arr.Find(d => d.dwReadingID == settings.DiskTempID && d.dwSensorIndex == settings.Disk3TempIndex)).Value)).ToString());


                //Disk speeds
                tmp2 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskReadID && d.dwSensorIndex == settings.Disk1SpeedID)).Value, 2);
                tmp3 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskWriteID && d.dwSensorIndex == settings.Disk1SpeedID)).Value, 2);
                values_to_send.Add(getSpeed(tmp2) + " " + getSpeed(tmp3));

                
                tmp2 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskReadID && d.dwSensorIndex == settings.Disk2SpeedID)).Value, 2);
                tmp3 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskWriteID && d.dwSensorIndex == settings.Disk2SpeedID)).Value, 2);
                values_to_send.Add(getSpeed(tmp2) + " " + getSpeed(tmp3));

                tmp2 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskReadID && d.dwSensorIndex == settings.Disk3SpeedID)).Value, 2);
                tmp3 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.DiskWriteID && d.dwSensorIndex == settings.Disk3SpeedID)).Value, 2);
                values_to_send.Add(getSpeed(tmp2) + " " + getSpeed(tmp3));

                //network
                tmp2 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.NetDownloadID && d.dwSensorIndex == settings.NetSpeedIndex)).Value, 2);
                values_to_send.Add(getNetSpeed(tmp2, true));

                tmp2 = Math.Round((data_arr.Find(d => d.dwReadingID == settings.NetUploadID && d.dwSensorIndex == settings.NetSpeedIndex)).Value, 2);
                values_to_send.Add(getNetSpeed(tmp2, false));


                resolve = string.Join(";", values_to_send.ToArray()) + "&";
                return resolve;
            } 
            catch (Exception ex)
            {
                eventLog.WriteEntry("Cant make data string: " + ex.Message, EventLogEntryType.Error);
            }
            return "";
        }


        private static string getStr(int val, string end)
        {
            if (val < 10) return val + end + "  ";
            else if (val < 100) return val + end + " ";
            else return val+end;
        }

        private static string getSpeed(double speed)
        {
            if (speed == 0) return "  0.00";
            else if (speed < 10) return "  " + speed.ToString("N2").Replace(",", ".");
            else if (speed < 100) return " " + speed.ToString("N2").Replace(",", ".");
            else if (speed > 999) return Math.Round(speed, 1).ToString("N2").Replace(",", ".");
            return speed.ToString("N2").Replace(",", ".");
        }

        private static string getNetSpeed(double speed, bool down)
        {
            if (speed < 10) return (down ? "DL:" : "UP:") + speed.ToString("N2").Replace(",", ".") + " KB/s";
            else if (speed < 99.5) return (down ? "DL:" : "UP:") + speed.ToString("0.0").Replace(",", ".") + " KB/s";
            else if (speed < 10000) return (down ? "DL:" : "UP:") + Math.Round((speed / 1024), 2).ToString("0.00").Replace(",", ".") + " MB/s";
            else if (speed >= 10000) return (down ? "DL:" : "UP:") + Math.Round((speed / 1024), 1).ToString("0.0").Replace(",", ".") + " MB/s";
            return (down ? "DL:" : "UP:") + ((int)Math.Round((speed / 1024), 0)).ToString().Replace(",", ".") + " MB/s";
        }

        private static string getRPM(int rpm)
        {
            if (rpm < 10) return "   " + rpm;
            else if (rpm < 100) return "  " + rpm;
            else if (rpm < 1000) return " " + rpm;
            else return "" + rpm;

        }


        private static void openSerialPort()
        {
            try
            {
                serialPort = new SerialPort(settings.SerialPortName, settings.SerialPortSpeed);
                serialPort.WriteTimeout = 1000;
                serialPort.Open();
                port_opened = true;
            } 
            catch (Exception ex)
            {
                port_opened = false;
                try { serialPort.Close(); } catch { }
                eventLog.WriteEntry("Cant open serial port: " + ex.Message, EventLogEntryType.Error);
            }
        }

        private static void timer_Tick(Object source, ElapsedEventArgs e)
        {
            if (port_opened)
            {
                if (HWiNFOisStarted())
                {
                    readMemory();
                    if (data_arr.Count > 0)
                    {
                        try
                        {
                            byte[] msg = Encoding.ASCII.GetBytes(makeSerialData());
                            serialPort.Write(msg, 0, msg.Length);
                        }
                        catch (Exception ex)
                        {
                            eventLog.WriteEntry("Sending data to serial fail: " + ex.Message, EventLogEntryType.Error);
                            port_opened = false;
                        }
                    }
                }
            } 
            else
            {
                openSerialPort();
            }
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct HWiNFO_SHARED_MEM
        {
            public uint dwSignature;
            public uint dwVersion;
            public uint dwRevision;
            public long poll_time;
            public uint dwOffsetOfSensorSection;
            public uint dwSizeOfSensorElement;
            public uint dwNumSensorElements;
            public uint dwOffsetOfReadingSection;
            public uint dwSizeOfReadingElement;
            public uint dwNumReadingElements;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct HWiNFO_ELEMENT
        {
            public SENSOR_READING_TYPE tReading;
            public uint dwSensorIndex;
            public uint dwReadingID;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = HWiNFO_SENSORS_STRING_LEN)]
            public string szLabelOrig;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = HWiNFO_SENSORS_STRING_LEN)]
            public string szLabelUser;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = HWiNFO_UNIT_STRING_LEN)]
            public string szUnit;
            public double Value;
            public double ValueMin;
            public double ValueMax;
            public double ValueAvg;
        }

        public enum SENSOR_READING_TYPE
        {
            SENSOR_TYPE_NONE,
            SENSOR_TYPE_TEMP,
            SENSOR_TYPE_VOLT,
            SENSOR_TYPE_FAN,
            SENSOR_TYPE_CURRENT,
            SENSOR_TYPE_POWER,
            SENSOR_TYPE_CLOCK,
            SENSOR_TYPE_USAGE,
            SENSOR_TYPE_OTHER,
        }
    }
}
