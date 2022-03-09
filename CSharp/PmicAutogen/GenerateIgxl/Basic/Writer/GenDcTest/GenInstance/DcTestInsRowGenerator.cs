using System;
using System.Collections.Generic;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenInstance
{
    public class DcTestInsRowGenerator : InsRowGenerator
    {
        public DcTestInsRowGenerator(string sheetName) : base(sheetName)
        {
        }

        public override List<InstanceRow> GenInsRows()
         {
            var instanceRows = new List<InstanceRow>();
            var insRow = GenHardIpRow();
            instanceRows.Add(insRow);
            return instanceRows;
        }

        protected InstanceRow GenHardIpRow()
        {
            var insRow = new InstanceRow();
            insRow.SheetName = SheetName;
            insRow.TestName = CreateHardIpTestName();
            insRow.Type = CreateType();
            insRow.Name = CreateVbtName();
            insRow.ArgList = CreateArgList();
            insRow.Args = CreateArgs();
            insRow.TimeSets = CreateDcTestTimeSets();
            insRow.DcCategory = CreateDcTestDcCategory();
            if (LabelVoltage.Equals("UHV",StringComparison.InvariantCultureIgnoreCase) ||
                LabelVoltage.Equals("ULV", StringComparison.InvariantCultureIgnoreCase))
            {
                insRow.DcCategory = LocalSpecs.GetUltraCategory(insRow.DcCategory);
            }
            insRow.DcSelector = CreateDcSelector();
            insRow.AcCategory = CreateDcTestAcCategory();
            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
            insRow.PinLevels = CreateDcTestPinLevel();
            return insRow;
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName);
            IpName = CommonGenerator.GetIpName(pattern.MiscInfo);
            SubBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo,
                BlockName, pattern.ForceCondition.IsCz2InstName);
            SubBlock2Name = CommonGenerator.GetSubBlock2Name(pattern.MiscInfo);
            TimingAc = CommonGenerator.GetTimingAc(pattern.AcUsed);
            InstNameSubStr = CommonGenerator.GetInstNameSubStr(pattern.MiscInfo);
            VbtFunction = CommonGenerator.GetVbtFunctionBase(pattern.FunctionName);
            if (VbtFunction.FunctionName == "" && pattern.MiscInfo != "")
                EpplusErrorManager.AddError(HardIpErrorType.MisVbtModule, ErrorLevel.Error, pattern.SheetName,
                    pattern.RowNum, "Can't find Vbt from MiscInfo: " + pattern.MiscInfo, "");
            NoPattern = CommonGenerator.NoPattern(pattern.Pattern.GetLastPayload());
            PowerOverWrite = CommonGenerator.GetHardIpDcSetting(pattern.LevelUsed);
        }

        protected string CreateHardIpTestName()
        {
            if (!string.IsNullOrEmpty(Pattern.TestName)) return Pattern.TestName + "_" + LabelVoltage;
            var patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, SubBlock2Name, IpName, patternName,
                Pattern.PatternIndexFlag, TimingAc, Pattern.DivideFlag, InstNameSubStr, LabelVoltage, NoPattern,
                Pat.WirelessData.IsNeedPostBurn, false, Pat.WirelessData.IsDoMeasure);
        }

        private string CreateDcTestDcCategory()
        {
            var dc = SearchInfo.GetSpecifyInfo(Pattern.DcCategory, "DC");
            if (!string.IsNullOrEmpty(dc))
                return dc;

            return HardIpConstData.LeakageDcDefault;
        }

        private string CreateDcTestAcCategory()
        {
            var ac = SearchInfo.GetSpecifyInfo(Pattern.AcCategory, "AC");
            if (!string.IsNullOrEmpty(ac))
                return ac;

            return HardIpConstData.AcCommonDefault;
        }

        private string CreateDcTestPinLevel()
        {
            if (!string.IsNullOrEmpty(Pattern.LevelUsed))
                return "Levels_" + Pattern.LevelUsed;

            return HardIpConstData.LeakageLevelDefault;
        }

        private string CreateDcTestTimeSets()
        {
            var patternName = Pattern.Pattern.GetLastPayload();
            if (InputFiles.PatternListMap.PatternListCsvRows.Exists(x =>
                x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase)))
            {
                var pattern = InputFiles.PatternListMap.PatternListCsvRows.Find(x =>
                    x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase));
                return pattern.ActualTimeSetVersion;
            }

            return string.Empty;
        }
    }
}