using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Library.Common;
using Library.DataStruct;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace Library.Input
{
    public class DatalogReader
    {
        public static Regex regDeviceNumber = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DeviceNumber").Pattern;
        public static Regex regInstanceLogHeader = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("InstanceLogHeader").Pattern;
        public static Regex regInstanceName = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("InstanceName").Pattern;
        public static Regex regInstanceLog = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("InstanceLog").Pattern;
        public static Regex regForceCondition = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("ForceCondition").Pattern;
        public static Regex regRestoreForceCondition = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("RestoreForceCondition").Pattern;
        public static Regex regSpecialForceCondition = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("SpecialForceCondition").Pattern;
        public static Regex regDigSrcStart = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DigSrcStart").Pattern;
        public static Regex regDigSrcEnd = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DigSrcEnd").Pattern;
        public static Regex regDigCapStart = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DigCapStart").Pattern;
        public static Regex regDigCapEnd = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DigCapEnd").Pattern;
        public static Regex regSrcBits = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("SrcBits").Pattern;
        public static Regex regSrcPin = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("SrcPin").Pattern;
        public static Regex regDataSequence = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DataSequence").Pattern;
        public static Regex regAssignment = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("Assignment").Pattern;
        public static Regex regDsscOut = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("DsscOut").Pattern;
        public static Regex regCapBits = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("CapBits").Pattern;
        public static Regex regCapPin = Common.CommonData.GetInstance().LogSettings.GetRowTypeTypePatternByName("CapPin").Pattern;
        public List<TestInstanceLogBase> ForTestNameLoglst = new List<TestInstanceLogBase>();
        private FileStream _fs;
        private StreamReader _sr;

        public DatalogReader()
        {

        }

        public List<Device> Read(string filePath)
        {
            List<Device> devicelist = new List<Device>();
            Device currentDevice = null;

            if (!File.Exists(filePath))
                return devicelist;

            OpenFile(filePath);

            TestLogType currentLogHeaderType = TestLogType.Undefined;
            TestInstanceItem currentInstance = null;
            string currentInstanceName = string.Empty;
            List<TestInstanceLogBase> currentInstanceLoglst = null;
            ForTestNameLoglst = new List<TestInstanceLogBase>();
            List<string> currentLogRowItems = null;
            TestInstanceLogBase currentTestLog = null;
            Regex currentLogReg = null;
            string currentDeviceNumber = string.Empty;

            try
            {
                int row = 1;
                string line;
                while ((line = _sr.ReadLine()) != null)
                {                                    
                    line = line.Trim();
                    DataLogRowType rowContextType = CheckLogRowType(line, currentDeviceNumber, currentLogHeaderType, currentLogReg);
                    switch (rowContextType)
                    {
                        case DataLogRowType.IgnoredRow:
                            break;
                        case DataLogRowType.Device:
                            currentDeviceNumber = regDeviceNumber.Match(line.Trim()).Groups["deviceNumber"].ToString();
                            currentDevice = new Device(currentDeviceNumber);
                            devicelist.Add(currentDevice);
                            break;
                        case DataLogRowType.Header:
                            HeaderPattern logHeaderSetting = CommonData.GetInstance().LogSettings.GetHeaderPatternByHeader(line);
                            if (logHeaderSetting.Name.Equals(TestInstanceLogMeasure.LogType, StringComparison.OrdinalIgnoreCase))
                            {
                                currentLogHeaderType = TestLogType.Measure;
                                currentLogReg = logHeaderSetting.DataRegex;
                            }
                            else
                            {
                                currentLogHeaderType = TestLogType.Undefined;
                                currentLogReg = null;
                            }
                            break;
                        case DataLogRowType.InstanceName:
                            currentInstanceName = regInstanceName.Match(line.Trim()).Groups["instanceName"].ToString();

                            if (currentInstance != null && currentInstance.IsValidInstance())
                            {
                                currentDevice.TestInstanceItemlist.Add(currentInstance);
                            }
                            currentInstanceLoglst = new List<TestInstanceLogBase>();
                            int insDuplicateIndex = currentDevice.TestInstanceItemlist.FindAll(s => s.InstanceName.Equals(currentInstanceName, StringComparison.OrdinalIgnoreCase)).ToList().Count;

                            currentInstance = new TestInstanceItem(currentInstanceName, currentDeviceNumber, row, insDuplicateIndex, currentInstanceLoglst);
                            break;
                        case DataLogRowType.InstanceLog:
                            if (currentLogHeaderType == TestLogType.Measure)
                            {                                
                                currentTestLog = new TestInstanceLogMeasure(row, line);
                                currentTestLog.InstanceName = currentInstanceName;
                                currentTestLog.DuplicateIndex = currentInstanceLoglst.FindAll(s => s.KeyWord.Equals(currentTestLog.KeyWord, StringComparison.OrdinalIgnoreCase)).ToList().Count;
                                currentInstanceLoglst.Add(currentTestLog);
                                ForTestNameLoglst.Add(currentTestLog);
                            }
                            break;
                    }                  
                    row++;
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error when reading datalog file. " + e.Message.ToString());
            }
            finally
            {
                CloseFile();
            }

            if (currentInstance != null && currentInstance.IsValidInstance())
            {
                currentDevice.TestInstanceItemlist.Add(currentInstance);
            }
            return devicelist;
        }

        private DataLogRowType CheckLogRowType(string lineContext,string currentDeviceNumber, TestLogType currentLogHeaderType, Regex currentLogReg)
        {
            if (currentLogHeaderType != TestLogType.Undefined)
            {
                if (currentLogReg.IsMatch(lineContext))
                    return DataLogRowType.InstanceLog;
            }

            if (regInstanceName.IsMatch(lineContext))
                return DataLogRowType.InstanceName;

            if (!string.IsNullOrEmpty(currentDeviceNumber) && regInstanceLogHeader.IsMatch(lineContext))
                return DataLogRowType.Header;
            if (regDeviceNumber.IsMatch(lineContext))
                return DataLogRowType.Device;

            return DataLogRowType.IgnoredRow;

        }

        private void OpenFile(string filePath)
        {
            _fs = new FileStream(filePath,FileMode.Open);
            _sr = new StreamReader(_fs);
        }

        private void CloseFile()
        {
            if (_sr != null)
            {
                _sr.Close();
                _sr.Dispose();
            }
            if (_fs != null)
            {
                _fs.Close();
                _fs.Dispose();
            }
        }
    }
}

public enum DataLogRowType
{
    Device,
    Header,
    InstanceName,
    InstanceLog,
    ForceCondition,
    RestoreForceCondition,
    SpecialForceCondition,
    DigSrcStart,
    DigSrcEnd,
    DigCapStart,
    DigCapEnd,
    SrcBits,
    SrcPin,
    DataSequence,
    Assignment,
    Dsscout,
    CapBits,
    CapPin,
    IgnoredRow
}
