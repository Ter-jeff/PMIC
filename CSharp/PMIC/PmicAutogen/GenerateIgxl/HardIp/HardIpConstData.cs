using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class HardIpConstData
    {
        public const string FlagClear = "Flag-Clear";
        public const string ReTestFlag = "F_Retest";
        public const string DdrLabel = "Func:DDR_EYE";
        public const string DigCapName = "MeasCapName";

        public const string FakePin = "FakePin";

        public static readonly List<string> LabelVolList = new List<string> {LabelNv, LabelLv, LabelHv};

        public static readonly List<string> MiscKeyList = new List<string>
        {
            Limit,
            Timing,
            Vbt,
            VbtKey,
            HighTemp,
            RoomTemp,
            InstanceName,
            RelayOn,
            RelayOff,
            NvOnly,
            HvOnly,
            LvOnly,
            RemoveHv,
            RemoveLv,
            RemoveNv,
            ReTest,
            FreeRunClkDisableWord,
            FreeRunClkEnableWord,
            NoBin,
            NoBinUseLimit,
            RemovePattern,
            Manual,
            SkipCheck,
            OpCode,
            InstNameSubStr,
            Calc,
            CalcParameter,
            DigCapName,
            ExecCond,
            SubBlockName,
            SubBlock2Name,
            SubBlockCzName,
            IpName,
            CallExtraFlow,
            RepeatLimit,
            Lsl,
            Usl,
            WiInitPat,
            WiPostPat,
            Slope,
            SweepRange,
            LoadFile,
            TransitionSlope,
            IgnorePatMeasC
        };

        # region default DC/AC/Timeset/Level

        public const string LeakageDcDefault = "Leakage";
        public const string AcCommonDefault = "Common";
        public const string LeakageLevelDefault = "Levels_Leakage";

        # endregion

        #region Misc info Standard name

        public const string Limit = "Limit";
        public const string VbtKey = "VBT";
        public const string Timing = "Timing";
        public const string Vbt = "Func|Trim|Meas";
        public const string OpCode = "opcode";
        public const string RegOpCode = @"opcode\s*:";
        public const string HighTemp = "High_Temp_Only";
        public const string RoomTemp = "Room_Temp_Only";
        public const string InstanceName = "InstanceName";
        public const string RelayOn = "RelayOn";
        public const string RelayOff = "RelayOff";

        public const string NvOnly = "NV_Only_For_HLN_Flow|NvOnly";
        public const string HvOnly = "HV_Only_For_HLN_Flow|HvOnly"; //Added to support HV_Only_For_HLN_Flow on 2016/6/27
        public const string LvOnly = "LV_Only_For_HLN_Flow|LvOnly"; //Added to support LV_Only_For_HLN_Flow on 2016/6/27
        public const string RemoveNv = @"Run_NV_Flow_Only|RemoveNv"; //Change to Run_NV_Flow_Only on 2016/6/28
        public const string RemoveLv = @"Run_LV_Flow_Only|RemoveLv"; //Change to Run_LV_Flow_Only on 2016/6/28
        public const string RemoveHv = @"Run_HV_Flow_Only|RemoveHv"; //Change to Run_HV_Flow_Only on 2016/6/28
        public const string ReTest = @"Fail_Retest|ReTest"; //Change from ReTest to Fail_Retest on 2016/6/23
        public const string PreNwireEnaOrDis = "FreerunClk";
        public const string FreeRunClkEnableWord = PreNwireEnaOrDis + "Enable";
        public const string FreeRunClkDisableWord = PreNwireEnaOrDis + "Disable";
        public const string NoBin = "No_Fail_Flag_All|NoBin"; //Change from NoBin to No_Fail_Flag_All on 2016/6/24

        public const string NoBinUseLimit = "No_Fail_Flag_UseLimit|NoBinUseLimit";
        public const string RemovePattern = @"Do_Not_Generate|RemovePattern";
        public const string Manual = @"Generate_But_Manually_Modify|Manual";
        public const string SkipCheck = "Skip_Pre_Check";
        public const string InstNameSubStr = "InstNameSubStr";
        public const string Calc = "Calc";
        public const string CalcParameter = "CalcArg";
        public const string IgnorePatInfo = "Ignore_Patt_Comment";
        public const string IgnorePatMeasC = "Ignore_Patt_MeasC";
        public const string RepeatLimit = "^Repeat_Limit";
        public const string NoPattern = "No_patt";
        public const string RegInsInPattern = @"^Instance[\s]*:[\s]*(?<InsName>[\w]+)";
        public const string Block = "Block";
        public const string IpName = "IP";
        public const string SubBlockName = "SubBlock";
        public const string SubBlock2Name = "SubBlock2";
        public const string SubBlockCzName = "SubBlockCZ";
        public const string ExecCond = "ExecCond";

        public const string CallExtraFlow = "Call";
        public const string KeepDsscOut = "Disable_MeasC_Split";
        public const string Cz2Only = "CZ2_Only";

        public const string Usl = "USL";
        public const string Lsl = "LSL";

        #endregion

        #region Naming rules

        public const string OpCodeTest = "Test";
        public const string OpCodeChar = "characterize";
        public const string OpCodeUseLimit = "Use-Limit";
        public const string OpCodeTestDeferLimit = "Test-defer-limits";
        public const string OpCodeBinTable = "Bintable";
        public const string HardipBinEnable = "HardIPBin";
        public const string EnvTtr = "TTR";
        public const string PrefixWireless = "WIRELESS_";
        public const string PrefixLcd = "LCD_";
        public const string PrefixDctest = "DCTEST_";
        public const string PrefixHardIpFailAction = "F";
        public const string SuffixHardIpFailAction = "_Flag";
        public const string TimeSetNa = "NA";
        public const string BinFlowFlag = "Bin";
        public const string PrefixInsSheetByVoltage = "TestInst_HARDIP_";
        public const string LabelAll = "All";
        public const string LabelNv = "NV";
        public const string LabelLv = "LV";
        public const string LabelHv = "HV";
        public const string LabelULv = "ULV";
        public const string LabelUHv = "UHV";

        public const string LabelHLv = "HLV";
        public const string SelectMax = "Max";
        public const string SelectMin = "Min";
        public const string SelectTyp = "Typ";
        public const string InstanceTypeDefault = "VBT";
        public const string PrefixReTest = "retest_";
        public const string RegCzPattern = @"cz_";
        public const string RegTestSequence = @"^TestSequence:";
        public const string SweepVoltage = @"sweep";
        public const string WiInitPat = "Wi_InitPat";
        public const string WiPostPat = "Wi_PostPat";
        public const string Slope = "Slope";
        public const string SweepRange = "SweepRange";
        public const string LoadFile = "LoadFile";
        public const string TransitionSlope = "TransitionSlope";

        #endregion

        #region Shmoo

        public const string PrefixShmooSetupName = "CZ2_";
        public const string RegShmoo = @"(xshmoo|yshmoo|sweep|sweepy)\s*\((?<ShmooStr>[^)]+)+\)";

        #endregion
    }
}