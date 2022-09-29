using CommonLib.Extension;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.GenDcTest.Writer
{
    internal class DcTestWriter
    {
        public const string BinFlowFlag = "Bin";
        public const string PrefixHardIpFailAction = "F";
        public const string SuffixHardIpFailAction = "_Flag";

        public const string All = "All";
        public const string Hnlv = "HNLV";
        public const string Nlv = "NLV";
        public const string Hlv = "HLV";
        public const string Hnv = "HNV";
        public const string Lv = "LV";
        public const string Nv = "NV";
        public const string Hv = "HV";
        public const string ULv = "ULV";
        public const string UHv = "UHV";
        public const string Max = "Max";
        public const string Min = "Min";
        public const string Typ = "Typ";
        public const string DcDefault = "Scan";
        public const string AcDefault = "Common";
        public const string LevelDefault = "Levels_Func";

        public readonly List<string> LabelVoltages = new List<string> { Nv, Lv, Hv };

        public DcTestWriter(string sheetName, List<HardIpPattern> patternList)
        {
            SheetName = sheetName;
            PatternList = patternList;
        }

        public string SheetName { get; }
        public List<HardIpPattern> PatternList { get; }

        protected string CreateBinTableName(HardIpPattern pattern, string voltage = "")
        {
            var patternName = pattern.Pattern.GetPatternName();
            if (Regex.IsMatch(patternName, RegInsInPattern, RegexOptions.IgnoreCase))
                patternName = Regex.Match(patternName, RegInsInPattern, RegexOptions.IgnoreCase)
                    .Groups["InsName"].ToString();

            var prefixSheetNameWithoutUnderscore = "_" + pattern.SheetName.Split('_')[0];
            var prefixBlockName = "_" + pattern.BlockName;
            var prefixSubBlock = string.IsNullOrEmpty(pattern.SubBlockName) ? string.Empty : "_" + pattern.SubBlockName;
            var prefixPatternName = patternName.GetSortPatNameForBinTable();
            if (prefixPatternName != "")
                prefixPatternName = "_" + prefixPatternName;
            var prefixTimingAc = string.IsNullOrEmpty(pattern.TimingAc) ? string.Empty : "_" + pattern.TimingAc;
            var prefixInstNameSubStr = string.IsNullOrEmpty(pattern.InstNameSubStr)
                ? string.Empty
                : "_" + pattern.InstNameSubStr;
            var prefixVoltage = voltage != "" ? "_" + voltage : "";
            if (pattern.NoPattern && !string.IsNullOrEmpty(pattern.InstNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            var parameter = BinFlowFlag + prefixSheetNameWithoutUnderscore + prefixBlockName +
                            GetSubBlockNameWithoutMinus(prefixSubBlock) + prefixPatternName + prefixTimingAc +
                            prefixInstNameSubStr + prefixVoltage;
            return parameter;
        }

        protected string CreateFailFlag(HardIpPattern pattern, string voltage)
        {
            var patternName = pattern.Pattern.GetPatternName();
            var prefixSheetNameWithoutUnderscore = "_" + pattern.SheetName.Split('_')[0];
            var prefixBlockName = "_" + pattern.BlockName;
            var prefixSubBlock = string.IsNullOrEmpty(pattern.SubBlockName) ? string.Empty : "_" + pattern.SubBlockName;
            var prefixSubBlock2 =
                string.IsNullOrEmpty(pattern.SubBlock2Name) ? string.Empty : "_" + pattern.SubBlock2Name;

            var prefixTimingAc = string.IsNullOrEmpty(pattern.TimingAc) ? string.Empty : "_" + pattern.TimingAc;
            var prefixLabelVoltage = string.IsNullOrEmpty(voltage) ? "" : "_" + voltage[0];
            if (prefixLabelVoltage != "" && voltage.StartsWith("U", StringComparison.OrdinalIgnoreCase))
                prefixLabelVoltage = "_" + voltage[0] + voltage[1];
            if (Regex.IsMatch(patternName, RegInsInPattern, RegexOptions.IgnoreCase))
                patternName = Regex.Match(patternName, RegInsInPattern, RegexOptions.IgnoreCase)
                    .Groups["InsName"].ToString();
            var prefixPatternName = patternName.GetSortPatNameForBinTable();
            if (prefixPatternName != "") prefixPatternName = "_" + prefixPatternName;
            var prefixInstNameSubStr =
                string.IsNullOrEmpty(pattern.InstNameSubStr) ? string.Empty : "_" + pattern.InstNameSubStr;
            if (pattern.NoPattern && !string.IsNullOrEmpty(pattern.InstNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            var failFlag = PrefixHardIpFailAction + prefixSheetNameWithoutUnderscore + prefixBlockName +
                           GetSubBlockNameWithoutMinus(prefixSubBlock) + prefixSubBlock2 + prefixPatternName +
                           prefixTimingAc + prefixInstNameSubStr + prefixLabelVoltage +
                           SuffixHardIpFailAction;

            if (pattern.MiscInfo.IsNoBin())
                failFlag = GetFlagNoBinStr(pattern.MiscInfo, voltage) + failFlag;
            return failFlag;
        }

        protected string CreateTestName(HardIpPattern pattern, string voltage)
        {
            if (!string.IsNullOrEmpty(pattern.TestName)) return pattern.TestName + "_" + voltage;
            var patternName = pattern.Pattern.GetPatternName();

            var blockName = pattern.BlockName;
            var subBlockName = pattern.SubBlockName;
            var subBlock2Name = pattern.SubBlock2Name;
            var prefixPatIndexFlag = pattern.PatternIndexFlag;
            var timingAc = pattern.TimingAc;
            var prefixDivideFlag = pattern.DivideFlag;
            var instNameSubStr = pattern.InstNameSubStr;
            var noPattern = pattern.NoPattern;
            var isGenByFlow = pattern.WirelessData.IsNeedPostBurn;
            var isDoMeasure = pattern.WirelessData.IsDoMeasure;

            #region pattern by "Instance:"

            if (Regex.IsMatch(patternName, RegInsInPattern, RegexOptions.IgnoreCase))
            {
                var insName = Regex.Match(patternName, RegInsInPattern, RegexOptions.IgnoreCase)
                    .Groups["InsName"].ToString();
                if (!insName.StartsWith(blockName + "_", StringComparison.CurrentCultureIgnoreCase))
                    insName = blockName + "_" + insName;
                if (isGenByFlow) return insName + prefixPatIndexFlag + prefixDivideFlag;
                const string instanceBlock = "INSREMOV_";
                patternName = instanceBlock + insName;
                return patternName.ToUpper() + prefixPatIndexFlag + prefixDivideFlag;
            }

            #endregion

            string prefixSubBlock;
            string prefixSubBlock2;
            subBlockName = subBlockName.Replace("-", "");
            {
                prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
                prefixSubBlock2 = string.IsNullOrEmpty(subBlock2Name) ? string.Empty : "_" + subBlock2Name;
            }

            var prefixTimingAc = string.IsNullOrEmpty(timingAc) ? string.Empty : "_" + timingAc;
            var prefixInstNameSubStr = string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr;
            var prefixLabelVoltage = string.IsNullOrEmpty(voltage) ? "" : "_" + voltage;
            var prefixPatternName = "_" + patternName;
            var prefixPostBurn = false ? "_PostBurn" : isDoMeasure ? "_DoMeasure" : "";
            if (noPattern && !string.IsNullOrEmpty(instNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            //Add prefixSubBlock by Marx on 2017/05/24
            //add condition : if substring exist -> timing AC should be cleared
            if (prefixInstNameSubStr != "") prefixTimingAc = "";
            var testName = blockName + prefixSubBlock + prefixSubBlock2 + prefixPatternName + prefixPatIndexFlag +
                           prefixTimingAc + prefixDivideFlag + prefixInstNameSubStr + prefixPostBurn +
                           prefixLabelVoltage;
            return testName.ToUpper();
        }

        public string GetSpecifyInfo(string dcInfo, string key)
        {
            foreach (var info in dcInfo.Split(';'))
            {
                var dcArray = info.Split(':');
                if (dcArray.Length != 2)
                    continue;
                if (!dcArray[0].Equals(key, StringComparison.OrdinalIgnoreCase))
                    continue;
                return dcArray[1];
            }

            return "";
        }

        public string GetAcSelector(string voltage, string acInfo)
        {
            var acSelector = "";
            foreach (var info in acInfo.Split(';'))
            {
                var acArray = info.Split(':');
                if (acArray.Length != 3) continue;
                if (acArray[1].Equals(voltage, StringComparison.OrdinalIgnoreCase)) acSelector = acArray[2];
                if (acArray[1].Equals(All, StringComparison.OrdinalIgnoreCase))
                    acSelector = acArray[2];
            }

            if (acSelector.Equals(Max, StringComparison.OrdinalIgnoreCase))
                acSelector = Max;
            if (acSelector.Equals(Min, StringComparison.OrdinalIgnoreCase))
                acSelector = Min;
            if (acSelector.Equals(Typ, StringComparison.OrdinalIgnoreCase))
                acSelector = Typ;

            return acSelector;
        }

        public string GetDcSelector(string labelVoltage, string dcInfo)
        {
            var dcSelector = "";
            foreach (var info in dcInfo.Split(';'))
            {
                var dcArray = info.Split(':');
                if (dcArray.Length != 3) continue;
                if (dcArray[1].Equals(labelVoltage, StringComparison.OrdinalIgnoreCase)) dcSelector = dcArray[2];
                if (dcArray[1].Equals(All, StringComparison.OrdinalIgnoreCase))
                    dcSelector = dcArray[2];
            }

            if (dcSelector.Equals(Max, StringComparison.OrdinalIgnoreCase))
                dcSelector = Max;
            if (dcSelector.Equals(Min, StringComparison.OrdinalIgnoreCase))
                dcSelector = Min;
            if (dcSelector.Equals(Typ, StringComparison.OrdinalIgnoreCase))
                dcSelector = Typ;

            return dcSelector;
        }

        public string GetEnvFromPattern(HardIpPattern pattern, bool isBinTableUse = false)
        {
            if (isBinTableUse)
            {
                if (pattern.NoBinOutStr.Equals("All", StringComparison.OrdinalIgnoreCase))
                    return "NoBinOut";
                if (LabelVoltages.Exists(p =>
                        Regex.IsMatch(pattern.NoBinOutStr, p, RegexOptions.IgnoreCase)))
                    return "SpecialBinOut";
            }

            var patternName = pattern.Pattern.GetLastPayload();
            if (patternName.Equals(NoPattern, StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(pattern.Pattern.GetLastPayload(), RegInsInPattern,
                    RegexOptions.IgnoreCase))
                return "";
            if (!pattern.IsInTestPlan)
                return "MissPattInTestPlan";

            string status;
            var missInPatternList = InputFiles.PatternListMap.GetStatusInPatternList(patternName, out status);
            if (!missInPatternList)
                return status;
            //bool missInScgh = GetStatusInScgh(HardIpDataMain.ScghData, patternName, out status);
            //if (!missInScgh)
            //    return status;
            return "";
        }

        public string GetFlagNoBinStr(string miscInfo, string voltage)
        {
            foreach (var item in miscInfo.Split(';'))
                if (miscInfo.IsNoBin())
                {
                    var noBinVoltage = item.Split(':').ToList();
                    if (noBinVoltage.Count > 1 && Regex.IsMatch(noBinVoltage[1], voltage))
                        return "No";
                }

            //string noBinVoltages = Regex.Match(miscInfo, NoBin + @":(?<setting>)", RegexOptions.IgnoreCase).Groups["setting"].ToString();
            //if (Regex.IsMatch(noBinVoltages, voltage, RegexOptions.IgnoreCase))
            //return "No";
            return "";
        }

        public string GetSubBlockNameWithoutMinus(string subBlockName)
        {
            return subBlockName.Replace("-", "");
        }

        #region pattern

        public const string RegInsInPattern = @"^Instance[\s]*:[\s]*(?<InsName>[\w]+)";
        public const string NoPattern = "No_patt";

        #endregion
    }
}