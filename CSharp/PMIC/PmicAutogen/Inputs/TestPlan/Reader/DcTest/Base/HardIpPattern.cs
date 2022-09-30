using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class HardIpPattern
    {
        public const string TtrHeader = "TTR";
        public const string NoBinOutHeader = "NoBinout";
        public const string TestItemHeader = @"Test\s*Item";
        public const string StepHeader = "Step";
        public const string DescriptionHeader = "Description";
        public const string PatternHeader = @"Pattern[\s]*([\/]Instance)?";
        public const string ForceConditionHeader = "Condition";
        public const string ForceConditionCharHeader = @"Condition\(Char\)"; //optional
        public const string RegisterAssignmentHeader = "Register";
        public const string MiscInfoHeader = @"Misc.*Info";
        public const string MeasHeader = @"^Meas";
        public const string LoHeader = @"\s*Lo\s*Limit.*";
        public const string HiHeader = @"\s*Hi\s*Limit.*";

        public const string CommentHeader = "Comment";

        //Add new header pattern due to ADC autogen
        public const string AnalogSetupHeader = @"Analog\s*Instrument\s*Setup";

        public const string TestNameHeader = @"^Test_*\s*Name";

        //Analog Instrument Setup	Test Name	Discard and Selected Cycle
        public static readonly List<string> KnownHeaders = new List<string>
        {
            TtrHeader,
            NoBinOutHeader,
            TestItemHeader,
            StepHeader,
            DescriptionHeader,
            PatternHeader,
            ForceConditionHeader,
            ForceConditionCharHeader,
            RegisterAssignmentHeader,
            MiscInfoHeader,
            MeasHeader,
            LoHeader,
            HiHeader,
            CommentHeader,
            AnalogSetupHeader,
            TestNameHeader
        };

        private string _functionName;

        #region Constructor

        public HardIpPattern()
        {
            RowNum = 0;
            ColumnNum = 0;
            SheetName = "";
            Pattern = new PatternClass("");
            ForceCondition = new ForceClass();
            ExtraPattern = null;
            TtrStr = "";
            NoBinOutStr = "";
            RelaySetting = new Dictionary<string, string>();
            NewRelaySetting = new Dictionary<string, string>();
            DupIndex = 0;
            ConditionIndex = 0;
            IsInTestPlan = true;
            MeasPins = new List<MeasPin>();
            OriMeasPins = new List<MeasPin>();
            UseLimitsH = new List<MeasPin>();
            UseLimitsL = new List<MeasPin>();
            UseLimitsN = new List<MeasPin>();
            MiscInfo = "";
            MiscInfoInTestPlan = "";
            ForceConditionList = new List<ForceCondition>();
            ForceConditionInTestPlan = new List<ForceCondition>();
            LevelUsedInTestPlan = "";
            TimeSetUsed = new TimeSetUsed();
            AcUsed = "";
            AcSelectorUsed = "";
            DcCategory = "";
            AcCategory = "";
            DcSelectorUsed = "";
            PassOrFail = "";
            RegisterAssignment = "";
            DigSrcAssignment = "";
            DigSrcEquation = "";
            FunctionName = "";
            DivideFlag = "";
            PatternType = "";
            LevelUsed = "";
            AcUsed = "";
            SweepCodes = new Dictionary<string, List<SweepCode>>();
            Shmoo = new HardipCharSetup();
            CalcEqn = "";
            TestPlanSequences = new List<TestPlanSequence>();
            TestPlanSequencesRf = new List<TestPlanSequence>();
            SpecialMeasType = "";
            InterposePostTest = "";
            SweepVoltage = new Dictionary<string, List<string>>();
            FlowControlFlag = -1;
            IsFlowInsRepeat = false;
            MixSigUsed = "";
            UseDeferLimit = false;
            WirelessData = new WirelessData();
            CustomVbName = "";
            CustomIndex = 0;
            VbtTypes = new List<PlanType> { PlanType.Default };
            PatTestName = "";
            InstrumentSetup = "";
            IsNonHardIpBlock = false;
            RfInterPose = "";
        }

        #endregion

        #region Property

        public int RowNum { get; set; }
        public int ColumnNum { get; set; }
        public string PassOrFail { get; set; }
        public string SheetName { get; set; }

        public string FunctionName
        {
            set { _functionName = value; }
            get { return _functionName.ToLower(); }
        }

        public string DivideFlag { get; set; }
        public int DivideIndex { get; set; }
        public string PatternType { get; set; }
        public string MiscInfo { get; set; }
        public ForceClass ForceCondition { get; set; }
        public HardipCharSetup Shmoo { get; set; }
        public string ForceCondChar { get; set; }
        public AnalogSrc AnalogSrc { get; set; }
        public AnalogDigCap AnalogCap { get; set; }
        public string TestName { get; set; }
        public string PatTestName { get; set; }
        public List<MeasPin> OriMeasPins { get; set; }
        public List<MeasPin> MeasPins { get; set; }
        public List<MeasPin> UseLimitsH { get; set; }
        public List<MeasPin> UseLimitsL { get; set; }
        public List<MeasPin> UseLimitsN { get; set; }
        public List<ForceCondition> ForceConditionList { get; set; }

        public string TtrStr { get; set; }
        public string NoBinOutStr { get; set; }
        public Dictionary<string, string> RelaySetting { get; set; }
        public Dictionary<string, string> NewRelaySetting { get; set; }
        public PatternClass Pattern { get; set; }
        public HardIpPattern ExtraPattern { get; set; }
        public int DupIndex { get; set; }
        public int ConditionIndex { get; set; }
        public List<ForceCondition> ForceConditionInTestPlan { get; set; }
        public string MiscInfoInTestPlan { get; set; }
        public string LevelUsedInTestPlan { get; set; }
        public string AcUsedInTestPlan { get; set; }
        public string TimeSetUsedInTestPlan { get; set; }
        public Dictionary<string, List<string>> SweepVoltage { get; set; }
        public string RegisterAssignment { get; set; }
        public string DigSrcAssignment { get; set; }
        public string DigSrcEquation { get; set; }
        public string CalcEqn { get; set; }
        public string LevelUsed { get; set; }
        public string AcUsed { get; set; }
        public string AcSelectorUsed { get; set; }
        public string DcCategory { get; set; }
        public string AcCategory { get; set; }
        public string DcSelectorUsed { get; set; }
        public TimeSetUsed TimeSetUsed { get; set; }
        public Dictionary<string, List<SweepCode>> SweepCodes { get; set; }
        public List<TestPlanSequence> TestPlanSequences { get; set; }
        public List<TestPlanSequence> TestPlanSequencesRf { get; set; }
        public string SpecialMeasType { get; set; }
        public string InterposePostTest { get; set; }
        public bool IsInTestPlan { get; set; } //If it is false,the pattern came from scgh
        public int FlowControlFlag { get; set; }
        public bool IsFlowInsRepeat { get; set; }
        public string MixSigUsed { get; set; }
        public bool UseDeferLimit { get; set; }
        public string CustomVbName { get; set; }
        public int CustomIndex { get; set; }
        public WirelessData WirelessData { get; set; }
        public List<PlanType> VbtTypes { get; set; }
        public string InstrumentSetup { get; set; }
        public bool IsNonHardIpBlock { get; set; }
        public string RfInterPose { get; set; }

        public string BlockName
        {
            get { return MiscInfo.GetBlockName(SheetName); }
        }

        public string SubBlockName
        {
            get { return MiscInfo.GetSubBlockName(Pattern.GetLastPayload(), BlockName); }
        }

        public string SubBlock2Name
        {
            get { return MiscInfo.GetSubBlock2Name(); }
        }

        public string IpName
        {
            get { return MiscInfo.GetIpName(); }
        }

        public string TimingAc
        {
            get { return ""; }
        }

        public string InstNameSubStr
        {
            get { return MiscInfo.GetInstNameSubStr(); }
        }

        public bool NoPattern
        {
            get { return GetNoPattern(Pattern.RealPatternName); }
        }

        public bool EnableWord
        {
            get { return GetNoPattern(Pattern.RealPatternName); }
        }

        #endregion

        #region member function

        private bool GetNoPattern(string patternName)
        {
            return patternName.Equals("No_patt", StringComparison.OrdinalIgnoreCase);
        }

        private string GetSubBlockNameByPattern(string patternName, string blockName, bool isCheckScghItem = true)
        {
            var subBlocks = new List<string>();
            var patternSeg = patternName.Split('_').ToList();
            var siDmIndex = patternSeg.FindLastIndex(p =>
                p.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                p.Equals("DM", StringComparison.OrdinalIgnoreCase));
            if (siDmIndex != -1 && siDmIndex != patternSeg.Count - 1)
            {
                var subBlockSegments = patternSeg.GetRange(siDmIndex + 1, patternSeg.Count - siDmIndex - 1);
                foreach (var subBlockSeg in subBlockSegments)
                    if (!subBlockSeg.Equals(blockName, StringComparison.CurrentCultureIgnoreCase) && isCheckScghItem)
                        subBlocks.Add(subBlockSeg);
                    else
                        subBlocks.Add(subBlockSeg);
            }

            return string.Join("_", subBlocks);
        }

        private string GetBlockNameFromSheetName(string sheetName)
        {
            var arr = sheetName.Split('_').ToList();
            if (arr.Count > 1)
                arr.RemoveAt(0);
            return string.Join("", arr).Replace(" ", "").ToUpper();
        }

        public void Copy(HardIpPattern pattern)
        {
            SheetName = pattern.SheetName;
            RowNum = pattern.RowNum;
            ColumnNum = pattern.ColumnNum;
            Pattern = pattern.Pattern;
            ExtraPattern = pattern.ExtraPattern;
            ConditionIndex = pattern.ConditionIndex;
            TtrStr = pattern.TtrStr;
            NoBinOutStr = pattern.NoBinOutStr;
            RelaySetting = pattern.RelaySetting;
            NewRelaySetting = pattern.NewRelaySetting;
            DupIndex = pattern.DupIndex;
            IsInTestPlan = pattern.IsInTestPlan;
            MiscInfo = pattern.MiscInfo;
            ForceConditionList = pattern.ForceConditionList;
            RegisterAssignment = pattern.RegisterAssignment;
            DigSrcAssignment = pattern.DigSrcAssignment;
            DigSrcEquation = pattern.DigSrcEquation;
            FunctionName = pattern.FunctionName;
            DivideFlag = pattern.DivideFlag;
            PatternType = pattern.PatternType;
            ForceConditionInTestPlan = pattern.ForceConditionInTestPlan;
            MiscInfoInTestPlan = pattern.MiscInfoInTestPlan;
            AcUsedInTestPlan = pattern.AcUsedInTestPlan;
            LevelUsedInTestPlan = pattern.LevelUsedInTestPlan;
            TimeSetUsedInTestPlan = pattern.AcUsedInTestPlan;
            LevelUsed = pattern.LevelUsed;
            AcUsed = pattern.AcUsed;
            AcSelectorUsed = pattern.AcSelectorUsed;
            DcCategory = pattern.DcCategory;
            AcCategory = pattern.AcCategory;
            DcSelectorUsed = pattern.DcSelectorUsed;
            TimeSetUsed = pattern.TimeSetUsed;
            SweepCodes = pattern.SweepCodes;
            Shmoo = pattern.Shmoo;
            CalcEqn = pattern.CalcEqn;
            TestPlanSequences = pattern.TestPlanSequences;
            TestPlanSequencesRf = pattern.TestPlanSequencesRf;
            InterposePostTest = pattern.InterposePostTest;
            SweepVoltage = pattern.SweepVoltage;
            FlowControlFlag = pattern.FlowControlFlag;
            IsFlowInsRepeat = pattern.IsFlowInsRepeat;
            ForceCondition = pattern.ForceCondition;
            ForceCondChar = pattern.ForceCondChar;
            AnalogSrc = pattern.AnalogSrc;
            AnalogCap = pattern.AnalogCap;
            MixSigUsed = pattern.MixSigUsed;
            UseDeferLimit = pattern.UseDeferLimit;
            WirelessData = new WirelessData(pattern.WirelessData);
            TestName = pattern.TestName;
            VbtTypes = pattern.VbtTypes;
            PatTestName = pattern.PatTestName;
            InstrumentSetup = pattern.InstrumentSetup;
            IsNonHardIpBlock = pattern.IsNonHardIpBlock;
            RfInterPose = pattern.RfInterPose;
            OriMeasPins = pattern.OriMeasPins;
        }

        public string PatternIndexFlag
        {
            get
            {
                var flag = "";
                if (SheetName.ToUpper() == "HARDIP_PCM") //Special PCM request by CC on 2016/04/26
                    if (ForceConditionList.Count == 1)
                        return "_" + ForceConditionList[0].ForcePins[0].ForceValue.Replace(".", "p") + "v";
                if (DupIndex > 0)
                    flag += "_" + DupIndex;
                if (ConditionIndex > 0)
                    flag += "_" + ConditionIndex;
                return flag;
            }
            set { throw new NotImplementedException(); }
        }

        #endregion
    }

    [Serializable]
    public class WirelessData
    {
        public WirelessData()
        {
        }

        public WirelessData(WirelessData data)
        {
            TrimFuseName = data.TrimFuseName;
            TrimTarget = data.TrimTarget;
            TrimMeas = data.TrimMeas;
            TrimCalcEqn = data.TrimCalcEqn;
            TrimType = data.TrimType;
            RegisterAssignment = data.RegisterAssignment;
            IsNeedPostBurn = data.IsNeedPostBurn;
            IsDoMeasure = data.IsDoMeasure;
        }

        public string TrimFuseName { get; set; }
        public string TrimTarget { get; set; }
        public string TrimMeas { get; set; }
        public string TrimCalcEqn { get; set; }
        public string TrimType { get; set; }
        public string RegisterAssignment { get; set; }

        public bool IsNeedPostBurn { get; set; }
        public bool IsDoMeasure { get; set; }
    }

    public enum PlanType
    {
        Default,
        Trim,
        Rf,
        Bb
    }
}