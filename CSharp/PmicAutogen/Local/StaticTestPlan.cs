using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;

namespace PmicAutogen.Local
{
    public static class StaticTestPlan
    {
        public static IfoldPowerTableSheet IfoldPowerTableSheet;
        public static PowerOverWriteSheet PowerOverWriteSheet;
        public static VddLevelsSheet VddLevelsSheet;
        public static IoLevelsSheet IoLevelsSheet;
        public static BscanCharSheet BscanCharSheet;
        public static DcTestContinuitySheet DcTestContinuitySheet;
        public static AhbRegisterMapSheet AhbRegisterMapSheet;

        private static List<ChannelMapSheet> _channelMapSheets;
        public static PinMapSheet IoPinMapSheet { get; set; }
        public static IoPinGroupSheet IoPinGroupSheet { get; set; }

        public static List<ChannelMapSheet> ChannelMapSheets
        {
            get { return _channelMapSheets ?? (_channelMapSheets = new List<ChannelMapSheet>()); }
            set
            {
                if (_channelMapSheets == null) _channelMapSheets = new List<ChannelMapSheet>();
                _channelMapSheets = value;
            }
        }

        public static List<TestPlanSheet> DcTestSheets { get; set; }
        public static PmicLeakageSheet PmicLeakageSheet { get; set; }
        public static PmicIdsSheet PmicIdsSheet { get; set; }
        public static SubFlowSheet MainFlowSheet { get; set; }
        public static PortDefineSheet PortDefineSheet { get; set; }
        public static OTPSetupSheet OtpSetupSheet { get; set; }

        public static void AddSheets(TestPlanManager testPlanManager)
        {
            IoPinMapSheet = testPlanManager.IoPinMapSheet;
            IoPinGroupSheet = testPlanManager.IoPinGroupSheet;
            ChannelMapSheets = testPlanManager.ChannelMapSheets;
            IfoldPowerTableSheet = testPlanManager.IfoldPowerTableSheet;
            PowerOverWriteSheet = testPlanManager.PowerOverWriteSheet;
            VddLevelsSheet = testPlanManager.VddLevelsSheet;
            IoLevelsSheet = testPlanManager.IoLevelsSheet;
            DcTestContinuitySheet = testPlanManager.DcTestContinuitySheet;
            AhbRegisterMapSheet = testPlanManager.AhbRegisterMapSheet;

            MainFlowSheet = testPlanManager.MainFlowSheet;
            PmicIdsSheet = testPlanManager.PmicIdsSheet;
            PmicLeakageSheet = testPlanManager.PmicLeakageSheet;
            PortDefineSheet = testPlanManager.PortDefineSheet;
            DcTestSheets = testPlanManager.DcTestSheet;
            OtpSetupSheet = testPlanManager.OtpSetupSheet;
            BscanCharSheet = testPlanManager.BscanCharSheet;
        }
    }
}