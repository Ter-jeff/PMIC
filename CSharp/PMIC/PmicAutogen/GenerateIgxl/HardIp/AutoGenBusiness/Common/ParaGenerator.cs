using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.PowerOverWrite;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common
{
    public class CommonGenerator
    {
        public static string GetTimingAc(string acSpecified)
        {
            return SearchInfo.GetTimingsByAc(acSpecified).Aggregate("",
                (current, timing) => current + timing.Name + "_" + timing.SuffixAcSpecName + "_").Trim('_');
        }

        public static string GetHardipSheetName(string sheetName)
        {
            sheetName = sheetName.Replace(" ",
                "_"); //Replace SheetName "HardIP BlockName" to "HardIP_BlockName" to prevent space name
            var generalSheets = new List<string> {HardIpConstData.PrefixWireless, HardIpConstData.PrefixLcd};
            var regRemovePrefix = string.Format(@"^({0})(?<Block>\w+)", string.Join("|", generalSheets));
            var block = Regex.Match(sheetName, regRemovePrefix, RegexOptions.IgnoreCase).Groups["Block"].Value;
            if (!string.IsNullOrEmpty(block))
                return block;

            return sheetName;
        }

        public static string GetBlockNameFromSheetName(string sheetName)
        {
            var arr = sheetName.Split('_').ToList();
            if (arr.Count > 1)
                arr.RemoveAt(0);
            return string.Join("", arr).Replace(" ", "").ToUpper();
        }

        public static string GetBlockName(string miscInfo, string sheetName)
        {
            var blockName = GetIpName(miscInfo);

            if (string.IsNullOrEmpty(blockName)) blockName = GetBlockNameFromSheetName(sheetName);
            return blockName;
        }

        public static string GetIpName(string miscInfo)
        {
            var blockName = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 &&
                    assignArr[0].Equals(HardIpConstData.IpName, StringComparison.OrdinalIgnoreCase))
                    blockName = assignArr[1].Replace("_", "");
            }

            return blockName;
        }

        public static bool GetPatternFailFlag(string miscInfo)
        {
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length >= 2 && assignArr[0].Equals("NoBinOut", StringComparison.OrdinalIgnoreCase) &&
                    assignArr[1].Equals("Pat", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        public static bool GetUseLimitFailFlag(string miscInfo)
        {
            foreach (var assign in miscInfo.Split(';'))
                if (assign.Equals("NoBinOut", StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        //public static bool GetNoBinTable(string miscInfo)
        //{
        //    foreach (var assign in miscInfo.Split(';'))
        //        if (assign.Equals("NoBinOut", StringComparison.OrdinalIgnoreCase))
        //            return true;
        //    return false;
        //}

        public static string GetSubBlockName(string patternName, string miscInfo, string blockName,
            bool isShmooInChar = false)
        {
            var subBlockName = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 && assignArr[0].Trim()
                    .Equals(HardIpConstData.SubBlockName, StringComparison.OrdinalIgnoreCase))
                {
                    subBlockName = assignArr[1].Replace("_", "");
                    break;
                }
            }

            if (string.IsNullOrEmpty(subBlockName))
                subBlockName = GetSubBlockNameByPattern(patternName, blockName);

            if (isShmooInChar)
                subBlockName += "CZ";
            return subBlockName;
        }

        public static string GetSubBlockNameByPattern(string patternName, string blockName, bool isCheckScghItem = true)
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

        public static string GetSubBlockNameWithoutMinus(string subBlockName)
        {
            return subBlockName.Replace("-", "");
        }

        public static string GetSubBlockNameWithoutMinus(string patternName, string miscInfo, string blockName,
            bool isShmooInChar = false)
        {
            return GetSubBlockName(patternName, miscInfo, blockName, isShmooInChar); //.Replace("-", "");
        }

        public static string GetSubBlock2Name(string miscInfo)
        {
            var subBlock2Name = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 &&
                    assignArr[0].Equals(HardIpConstData.SubBlock2Name, StringComparison.OrdinalIgnoreCase))
                {
                    subBlock2Name = assignArr[1].Replace("_", "");
                    break;
                }
            }

            return subBlock2Name;
        }

        public static string GetInstNameSubStr(string miscInfo)
        {
            return SearchInfo.GetInstNameSubStr(miscInfo);
        }

        public static PowerOverWrite GetHardIpDcSetting(string levelUsed)
        {
            if (!string.IsNullOrEmpty(levelUsed) && HardIpDataMain.PowerOverWriteSheet != null)
            {
                var levelSetting = HardIpDataMain.PowerOverWriteSheet.PowerOverWrite.FirstOrDefault(a =>
                    a.CategoryName.Equals(levelUsed, StringComparison.OrdinalIgnoreCase));
                return levelSetting;
            }

            return null;
        }

        public static VbtFunctionBase GetVbtFunctionBase(string functionName)
        {
            return TestProgram.VbtFunctionLib.GetFunctionByName(functionName);
        }

        public static bool NoPattern(string patternName)
        {
            return patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsCz2Only(string miscInfo)
        {
            var math = miscInfo.Split(';').ToList()
                .Exists(s => s.Equals(HardIpConstData.Cz2Only, StringComparison.OrdinalIgnoreCase));
            return math;
        }

        public static bool HasRetest(string miscInfo)
        {
            return Regex.IsMatch(miscInfo, HardIpConstData.ReTest, RegexOptions.IgnoreCase);
        }

        public static bool HasSweepCode(HardIpPattern pattern)
        {
            return pattern.SweepCodes.Count > 0;
        }

        public static bool HasSweepVoltage(HardIpPattern pattern)
        {
            return pattern.SweepVoltage.Count > 0;
        }

        public static bool HasShmoo(HardIpPattern pattern)
        {
            return pattern.Shmoo.CharSteps.Count > 0;
        }

        public static string ActualLabelVoltage(string labelVoltage, HardIpPattern pattern)
        {
            var voltage = labelVoltage;
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.HvOnly, RegexOptions.IgnoreCase))
                voltage = HardIpConstData.LabelHv;
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.LvOnly, RegexOptions.IgnoreCase))
                voltage = HardIpConstData.LabelLv;
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.NvOnly, RegexOptions.IgnoreCase))
                voltage = HardIpConstData.LabelNv;
            return voltage;
        }

        public static bool PatFlowNoNeedToGen(string labelVoltage, string miscInfo)
        {
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveHv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveLv, RegexOptions.IgnoreCase))
                        return true;
                    break;
                case HardIpConstData.LabelLv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveHv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveNv, RegexOptions.IgnoreCase))
                        return true;
                    break;
                case HardIpConstData.LabelHv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveNv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveLv, RegexOptions.IgnoreCase))
                        return true;
                    break;
            }

            return false;
        }

        public static string GenHardIpInsTestName(string blockName, string subBlockName, string subBlock2Name,
            string ipName, string patternName, string prefixPatIndexFlag, string timingAc, string prefixDivideFlag,
            string instNameSubStr, string labelVoltage, bool noPattern, bool isPostBurn, bool isGenByFlow,
            bool isDoMeas)
        {
            #region pattern by "Instance:"

            if (Regex.IsMatch(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
            {
                var insName = Regex.Match(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase)
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
            if (HardIpDataMain.ConfigData.InstanceNamingRule.Equals("new", StringComparison.OrdinalIgnoreCase))
            {
                prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? "_X" : "_" + subBlockName;
                prefixSubBlock2 = string.IsNullOrEmpty(subBlock2Name) ? "_X" : "_" + subBlock2Name;
            }
            else
            {
                prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
                prefixSubBlock2 = string.IsNullOrEmpty(subBlock2Name) ? string.Empty : "_" + subBlock2Name;
            }

            var prefixTimingAc = string.IsNullOrEmpty(timingAc) ? string.Empty : "_" + timingAc;
            var prefixInstNameSubStr = string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr;
            var prefixLabelVoltage = string.IsNullOrEmpty(labelVoltage) ? "" : "_" + labelVoltage;
            var prefixPatternName = "_" + patternName;
            var prefixPostBurn = isPostBurn ? "_PostBurn" : isDoMeas ? "_DoMeasure" : "";
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

        public static string GenHardIpFlowTestFailAction(string sheetName, string blockName, string subBlockName,
            string subBlock2Name, string patternName, string timingAc, string instNameSubStr, string labelVoltage,
            string miscInfo, bool noPattern)
        {
            var failAction = GenHardIpFlowFailFlag(sheetName, blockName, subBlockName, subBlock2Name, patternName,
                timingAc, instNameSubStr, labelVoltage, noPattern);
            if (Regex.IsMatch(miscInfo, HardIpConstData.NoBin, RegexOptions.IgnoreCase))
                failAction = SearchInfo.GetFlagNoBinStr(miscInfo, labelVoltage) + failAction;
            return failAction;
        }

        public static string GenHardIpFlowUseLimitFailAction(string sheetName, string blockName, string subBlockName,
            string subBlock2Name, string patternName, string timingAc, string instNameSubStr, string labelVoltage,
            string miscInfo, bool noPattern)
        {
            var failAction = GenHardIpFlowFailFlag(sheetName, blockName, subBlockName, subBlock2Name, patternName,
                timingAc, instNameSubStr, labelVoltage, noPattern);
            if (Regex.IsMatch(miscInfo, HardIpConstData.NoBinUseLimit, RegexOptions.IgnoreCase))
                failAction = SearchInfo.GetFlagNoBinUseLimitStr(miscInfo, labelVoltage) + failAction;
            return failAction;
        }

        public static string GenHardIpFlowBinParameter(string sheetName, string blockName, string subBlockName,
            string patternName, string timingAc, string instNameSubStr, bool noPattern, string voltage = "")
        {
            if (Regex.IsMatch(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
                patternName = Regex.Match(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase)
                    .Groups["InsName"].ToString();

            var prefixSheetNameWithoutUnderscore = "_" + sheetName.Split('_')[0];
            var prefixBlockName = "_" + blockName;
            var prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
            var prefixPatternName = patternName != "" ? "_" + patternName : "";
            prefixPatternName = patternName.GetSortPatNameForBinTable();
            if (prefixPatternName != "") prefixPatternName = "_" + prefixPatternName;
            var prefixTimingAc = string.IsNullOrEmpty(timingAc) ? string.Empty : "_" + timingAc;
            var prefixInstNameSubStr =
                string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr; //20180824 add
            var prefixVoltage = voltage != "" ? "_" + voltage : "";
            if (noPattern && !string.IsNullOrEmpty(instNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            var parameter = HardIpConstData.BinFlowFlag + prefixSheetNameWithoutUnderscore + prefixBlockName +
                            GetSubBlockNameWithoutMinus(prefixSubBlock) + prefixPatternName + prefixTimingAc +
                            prefixInstNameSubStr + prefixVoltage;
            return parameter;
        }

        public static string GenHardIpFlowBinParameter(string sheetName, string blockName, string subBlockName)
        {
            var binItemsList = new List<string>();
            binItemsList.Add(blockName);
            binItemsList.Add(subBlockName);
            if (sheetName.ToLower().Contains("hardip_"))
                return ComCombine.CombineByUnderLine("HIP", string.Join("_", binItemsList));
            return string.Join("_", binItemsList);
        }

        public static string GenEnableWord(string patternName, string miscInfo, string labelVoltage)
        {
            const string czPatEnableW = "_CZ";
            const string prefixEnable = "HardIP_";

            if (string.IsNullOrEmpty(labelVoltage))
                return "";

            var isCzPattern = Regex.IsMatch(patternName, HardIpConstData.RegCzPattern, RegexOptions.IgnoreCase);
            var enableWord = prefixEnable + labelVoltage;
            if (isCzPattern)
                enableWord += czPatEnableW;
            if (NeedRemoveEnableWord(miscInfo, labelVoltage))
                enableWord = "";
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (!HardIpDataMain.TestPlanData.NvEnable && !isCzPattern)
                        enableWord = "";
                    if (!HardIpDataMain.TestPlanData.CzNvEnable && isCzPattern)
                        enableWord = "";
                    break;
                case HardIpConstData.LabelLv:
                    if (!HardIpDataMain.TestPlanData.LvEnable && !isCzPattern)
                        enableWord = "";
                    if (!HardIpDataMain.TestPlanData.CzLvEnable && isCzPattern)
                        enableWord = "";
                    break;
                case HardIpConstData.LabelHv:
                    if (!HardIpDataMain.TestPlanData.HvEnable && !isCzPattern)
                        enableWord = "";
                    if (!HardIpDataMain.TestPlanData.CzHvEnable && isCzPattern)
                        enableWord = "";
                    break;
            }

            return enableWord;
        }

        private static bool NeedRemoveEnableWord(string miscInfo, string labelVoltage)
        {
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveNv, RegexOptions.IgnoreCase))
                        return true;
                    break;
                case HardIpConstData.LabelHv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveHv, RegexOptions.IgnoreCase))
                        return true;
                    break;
                case HardIpConstData.LabelLv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RemoveLv, RegexOptions.IgnoreCase))
                        return true;
                    break;
            }

            return false;
        }

        public static string GenHardIpFlowFailFlag(string sheetName, string blockName, string subBlockName,
            string subBlock2Name, string patternName, string timingAc, string instNameSubStr, string labelVoltage,
            bool noPattern)
        {
            var prefixSheetNameWithoutUnderscore = "_" + sheetName.Split('_')[0];
            var prefixBlockName = "_" + blockName;
            string prefixSubBlock;
            string prefixSubBlock2;
            if (HardIpDataMain.ConfigData.InstanceNamingRule.Equals("new", StringComparison.OrdinalIgnoreCase))
            {
                prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? "_X" : "_" + subBlockName;
                prefixSubBlock2 = string.IsNullOrEmpty(subBlock2Name) ? "_X" : "_" + subBlock2Name;
            }
            else
            {
                prefixSubBlock = string.IsNullOrEmpty(subBlockName) ? string.Empty : "_" + subBlockName;
                prefixSubBlock2 = string.IsNullOrEmpty(subBlock2Name) ? string.Empty : "_" + subBlock2Name;
            }

            var prefixTimingAc = string.IsNullOrEmpty(timingAc) ? string.Empty : "_" + timingAc;
            var prefixLabelVoltage = string.IsNullOrEmpty(labelVoltage) ? "" : "_" + labelVoltage[0];
            if (prefixLabelVoltage != "" && labelVoltage.StartsWith("U",StringComparison.OrdinalIgnoreCase))
            {
                prefixLabelVoltage = "_" + labelVoltage[0] + labelVoltage[1];
            }
            if (Regex.IsMatch(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
                patternName = Regex.Match(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase)
                    .Groups["InsName"].ToString();
            var prefixPatternName = "_" + patternName;
            prefixPatternName = patternName.GetSortPatNameForBinTable();
            if (prefixPatternName != "") prefixPatternName = "_" + prefixPatternName;
            var prefixInstNameSubStr =
                string.IsNullOrEmpty(instNameSubStr) ? string.Empty : "_" + instNameSubStr; //20180824 add
            if (noPattern && !string.IsNullOrEmpty(instNameSubStr))
            {
                prefixPatternName = prefixInstNameSubStr;
                prefixInstNameSubStr = "";
            }

            var failFlag = HardIpConstData.PrefixHardIpFailAction + prefixSheetNameWithoutUnderscore + prefixBlockName +
                           GetSubBlockNameWithoutMinus(prefixSubBlock) + prefixSubBlock2 + prefixPatternName +
                           prefixTimingAc + prefixInstNameSubStr + prefixLabelVoltage +
                           HardIpConstData.SuffixHardIpFailAction;
            return failFlag;
        }

        public static string GetRepeatMapping(string miscInfo)
        {
            foreach (var param in miscInfo.Split(';'))
            {
                if (!param.Contains(":"))
                    continue;
                var paramName = param.Split(':')[0];
                var paramValue = param.Split(':')[1];
                if (paramName == HardIpConstData.Limit)
                    return paramValue;
            }

            return "";
        }

        public static string CalculateLimit(string limitStr, string repeatStr)
        {
            if (limitStr == "" || limitStr == "0")
                return limitStr;

            const string regRate = @"(?<Num>\d+(\.\d){0,1})x+$";
            var limitValue = Convert.ToDouble(limitStr);
            double rate = 1;
            if (repeatStr != "")
                rate = double.Parse(Regex.Match(repeatStr, regRate).Groups["Num"].ToString());
            var result = limitValue * rate;
            return result.ToString(CultureInfo.InvariantCulture);
        }

        public static void ConvertPatNameInOpCode(List<string> opCodeList, FlowRows flowRowsBefore, string voltage)
        {
            if (opCodeList == null || opCodeList.Count == 0)
                return;
            var suffixFailAction = voltage[0] + HardIpConstData.SuffixHardIpFailAction;
            for (var i = 0; i < opCodeList.Count; i++)
            {
                var opCode = opCodeList[i].Split(':');
                if (opCode.Length != 2)
                    continue;

                var reg = new Regex(@"[\w]+");
                var newOpCode = reg.Replace(opCode[1], delegate(Match m)
                {
                    var patternName = m.Value;
                    if (SearchInfo.IsValidPatName(patternName))
                    {
                        var failFlagList = flowRowsBefore
                            .Where(s => Regex.IsMatch(s.Parameter, patternName, RegexOptions.IgnoreCase) &&
                                        s.OpCode != HardIpConstData.OpCodeUseLimit &&
                                        s.FailAction.EndsWith(suffixFailAction)).Select(s => s.FailAction).ToList();
                        if (failFlagList.Count > 0) return string.Join("||", failFlagList);
                    }

                    return patternName;
                });
                opCode[1] = newOpCode;
                opCodeList[i] = string.Join(":", opCode.ToList());
            }
        }
    }
}