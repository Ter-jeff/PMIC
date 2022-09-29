using System.Collections.Generic;
using System.Linq;
using AutomationCommon.EpplusErrorReport;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData
{
    public class HardIpDataMain
    {
        public static void Initialize(bool isReadSetting = true)
        {
            #region Initial data structure

            TestPlanData = new TestPlanData();
            ConfigData = new ConfigData();
            //ScghData = new ScghData();
            PatInfoData = new PatInfoData();
            PowerOverWriteSheet = new PowerOverWriteSheet();
            NewGenPatSet = new List<string>();
            //ScanConfigFilePath = LocalSpecs.SettingFolder + string.Format("\\Settings\\SCGH\\HardIP_Config_{0}.xml", LocalSpecs.CurrentProject);
            if (isReadSetting)
            {
                //if (LocalSpecs.Device != DeviceEnum.LCD)
                //    NwirePinsList = NwireSingleton.Instance().SettingInfo.NwirePins;

                //ShmooParameterTypeDictionary = HardipCharSetup.GetShmooParameterTypeDictionary();
            }

            //HardIpRegAssignDictionary = new List<HardIpRegAssign>();
            //HardipDcDefault = new PowerTableSheet();
            //MultiTestSettingSheets = new MultiTestSettingSheets(null);
            CorePwrFromBc = new List<string>();

            #endregion
        }

        public static void ReadPinMap()
        {
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null)
            {
                const string outString = "Missing IO_PinMap sheet in TestPlan";
                EpplusErrorManager.AddError(HardIpErrorType.MissingNeededSheets.ToString(), ErrorLevel.Error, "", 1,
                    outString);
                return;
            }

            foreach (var pin in pinMap.PinList)
                if (!TestPlanData.PinList.ContainsKey(pin.PinName.ToUpper()))
                    TestPlanData.PinList.Add(pin.PinName.ToUpper(), pin.PinType);

            foreach (var pinGroup in pinMap.GroupList)
            {
                if (!TestPlanData.PinGroupList.ContainsKey(pinGroup.PinName.ToUpper()))
                {
                    TestPlanData.PinGroupList.Add(pinGroup.PinName.ToUpper(),
                        pinGroup.PinList.Select(x => x.PinName).ToList());
                }
                else
                {
                    var outString = "Duplicate pin groups --- " + pinGroup.PinName;
                    EpplusErrorManager.AddError(HardIpErrorType.MissingNeededSheets.ToString(), ErrorLevel.Error, "", 1,
                        outString);
                    continue;
                }

                if (!TestPlanData.PinList.ContainsKey(pinGroup.PinName.ToUpper()))
                    TestPlanData.PinList.Add(pinGroup.PinName.ToUpper(), pinGroup.PinType);
            }
        }

        #region Universal static Data for HardIP part

        public static TestPlanData TestPlanData { get; set; }

        public static ConfigData ConfigData { get; set; }

        //public static ScghData ScghData { get; set; }
        public static PatInfoData PatInfoData { get; set; }

        public static List<string> NewGenPatSet { get; set; }
        //public static List<ProtocolAwarePin> NwirePinsList;

        public static PowerOverWriteSheet PowerOverWriteSheet;

        public static List<string> CorePwrFromBc { get; set; }

        #endregion
    }
}