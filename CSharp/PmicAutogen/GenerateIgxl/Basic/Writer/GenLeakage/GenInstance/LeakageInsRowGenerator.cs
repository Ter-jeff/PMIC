using System;
using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenInstance
{
    public class LeakageInsRowGenerator : InsRowGenerator
    {
        #region Constructor

        public LeakageInsRowGenerator(string sheetName) : base(sheetName)
        {
        }

        #endregion

        #region Main Methods

        public override List<InstanceRow> GenInsRows()
        {
            var instanceRows = new List<InstanceRow>();
            var insRow = new InstanceRow();
            insRow.SheetName = SheetName;
            insRow.Type = CreateType();
            insRow.Name = CreateVbtName();
            insRow.ArgList = CreateArgList();
            insRow.Args = CreateArgs();
            insRow.TestName = CreateLeakageTestName();
            insRow.DcCategory = CreateLeakageDcCategory();
            insRow.DcSelector = CreateDcSelector();
            insRow.AcCategory = CreateLeakageAcCategory();
            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector();
            insRow.PinLevels = !string.IsNullOrEmpty(Pat.LevelUsed)
                ? "Levels_" + Pat.LevelUsed
                : CreateLeakagePinLevel();
            insRow.TimeSets = CreateLeakageTimeSets();
            instanceRows.Add(insRow);
            return instanceRows;
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName);
            SubBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo,
                BlockName, pattern.ForceCondition.IsCz2InstName);
            SubBlock2Name = CommonGenerator.GetSubBlock2Name(pattern.MiscInfo);
            IpName = CommonGenerator.GetIpName(pattern.MiscInfo);
            TimingAc = CommonGenerator.GetTimingAc(pattern.AcUsed);
            InstNameSubStr = CommonGenerator.GetInstNameSubStr(pattern.MiscInfo);
            VbtFunction = CommonGenerator.GetVbtFunctionBase(pattern.FunctionName);
            NoPattern = CommonGenerator.NoPattern(pattern.Pattern.GetLastPayload());
            PowerOverWrite = CommonGenerator.GetHardIpDcSetting(pattern.LevelUsed);
        }

        #endregion

        #region Generate each columns Methods

        private string CreateLeakageTestName()
        {
            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, SubBlock2Name, IpName,
                Pattern.Pattern.GetLastPayload(), Pattern.PatternIndexFlag, TimingAc, Pattern.DivideFlag,
                InstNameSubStr, LabelVoltage, NoPattern, Pat.WirelessData.IsNeedPostBurn, false,
                Pat.WirelessData.IsDoMeasure);
        }

        private string CreateLeakageDcCategory()
        {
            var dc = SearchInfo.GetSpecifyInfo(Pattern.DcCategory, "DC");
            if (!string.IsNullOrEmpty(dc))
                return dc;

            return HardIpConstData.LeakageDcDefault;
        }

        private string CreateLeakageAcCategory()
        {
            var ac = SearchInfo.GetSpecifyInfo(Pattern.AcCategory, "AC");
            if (!string.IsNullOrEmpty(ac))
                return ac;

            return HardIpConstData.AcCommonDefault;
        }

        private string CreateLeakagePinLevel()
        {
            if (!string.IsNullOrEmpty(Pattern.LevelUsed))
                return "Levels_" + Pattern.LevelUsed;

            return HardIpConstData.LeakageLevelDefault;
        }

        private string CreateLeakageTimeSets()
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

        #endregion
    }
}