using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using AutomationCommon.EpplusErrorReport;
using IgxlData.NonIgxlSheets;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility
{
    public class SearchInfo
    {
        public static HardIpReference GetHardIpInfo(string patternName)
        {
            var hardIpInfo = new HardIpReference();
            if (HardIpDataMain.PatInfoData == null)
                return hardIpInfo;
            foreach (var info in HardIpDataMain.PatInfoData.PatInfoList)
                if (info.Payload.Equals(patternName, StringComparison.OrdinalIgnoreCase))
                {
                    hardIpInfo = info;
                    break;
                }

            return hardIpInfo;
        }

        public static HardIpReference GetHardIpInfo(HardIpPattern pattern)
        {
            var hardIpInfo = GetHardIpInfo(pattern.Pattern.GetLastPayload());
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase) &&
                hardIpInfo.SeqInfo.Count > 0)
            {
                var newHardIpInfo = new HardIpReference();
                newHardIpInfo.CapBit = hardIpInfo.CapBit;
                newHardIpInfo.CapBitName = hardIpInfo.CapBitName;
                newHardIpInfo.CapBitStr = hardIpInfo.CapBitStr;
                newHardIpInfo.DigSrcAssign = hardIpInfo.DigSrcAssign;
                newHardIpInfo.DigSrcDataWidth = hardIpInfo.DigSrcDataWidth;
                newHardIpInfo.DsscOut = hardIpInfo.DsscOut;
                newHardIpInfo.Payload = hardIpInfo.Payload;
                newHardIpInfo.SendBit = hardIpInfo.SendBit;
                newHardIpInfo.SendBitName = hardIpInfo.SendBitName;
                newHardIpInfo.SendBitStr = hardIpInfo.SendBitStr;
                newHardIpInfo.SendPinName = hardIpInfo.SendPinName;
                newHardIpInfo.SubRoutine = hardIpInfo.SubRoutine;
                newHardIpInfo.TimeSet = hardIpInfo.TimeSet;
                newHardIpInfo.IsIgnoreComment = true;
                return newHardIpInfo;
            }

            return hardIpInfo;
        }

        public static Dictionary<string, string> GetTestLimitPerMeasType(HardIpPattern pattern)
        {
            var dicLimitPerPin = new Dictionary<string, string>();
            var patInfo = GetHardIpInfo(pattern);
            var seqCount = patInfo == null || patInfo.SeqInfo.Count == 0
                ? pattern.TestPlanSequences.Count
                : patInfo.SeqInfo.Count;

            #region Gets TestLimitPerPin_VFI for vif

            if (Regex.IsMatch(pattern.FunctionName, VbtFunctionLib.VifName + "|" + VbtFunctionLib.FreqSynMeasFreqCurr,
                RegexOptions.IgnoreCase))
            {
                dicLimitPerPin.Add("V", "F");
                dicLimitPerPin.Add("F", "F");
                dicLimitPerPin.Add("I", "F");

                #region exists Sequence in patInfo

                for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    var pinList = new List<string>();
                    var measType = "";
                    foreach (var measPin in pattern.MeasPins)
                        if (measPin.SequenceIndex == sequenceIndex && !Regex.IsMatch(measPin.MeasType,
                            @"Calc|Limit|MeasC", RegexOptions.IgnoreCase))
                        {
                            pinList.Add(measPin.PinName);
                            measType = measPin.MeasType;
                        }

                    if (pinList.Count == 1 && !pinList[0].Contains(",") && !pinList[0].Contains("::"))
                        continue;
                    if (measType.Equals("MeasV", StringComparison.OrdinalIgnoreCase) ||
                        measType.Equals("MeasVdiff", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["V"] = "T";
                    if (measType.Equals("MeasI", StringComparison.OrdinalIgnoreCase) ||
                        measType.Equals("MeasI2", StringComparison.OrdinalIgnoreCase) ||
                        measType.Equals("MeasIdiff", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["I"] = "T";
                    if (measType.Equals("MeasF", StringComparison.OrdinalIgnoreCase) ||
                        measType.Equals("MeasFdiff", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["F"] = "T";
                }

                return dicLimitPerPin;

                #endregion
            }

            #endregion

            #region Gets TestLimitPerPin_VIR for vir

            if (Regex.IsMatch(pattern.FunctionName, VbtFunctionLib.VirName, RegexOptions.IgnoreCase))
            {
                dicLimitPerPin.Add("V", "F");
                dicLimitPerPin.Add("I", "F");
                dicLimitPerPin.Add("R", "F");

                #region exists Sequence in patInfo

                for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    var pinList = new List<string>();
                    var measType = "";
                    foreach (var measPin in pattern.MeasPins)
                        if (measPin.SequenceIndex == sequenceIndex && !Regex.IsMatch(measPin.MeasType,
                            @"Calc|Limit|MeasC", RegexOptions.IgnoreCase))
                        {
                            pinList.Add(measPin.PinName);
                            measType = measPin.MeasType;
                        }

                    if (pinList.Count == 1 && !pinList[0].Contains(",") && !pinList[0].Contains("::"))
                        continue;
                    if (measType.Equals("MeasV", StringComparison.OrdinalIgnoreCase) ||
                        measType.Equals("MeasE", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["V"] = "T";
                    if (measType.Equals("MeasI", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["I"] = "T";
                    if (measType.Equals("MeasR1", StringComparison.OrdinalIgnoreCase)) dicLimitPerPin["R"] = "T";
                }

                return dicLimitPerPin;

                #endregion
            }

            #endregion

            return dicLimitPerPin;
        }

        public static bool GetFlagSingleLimit(HardIpPattern pattern, string labelVoltage)
        {
            var patInfo = GetHardIpInfo(pattern);
            var seqCount = patInfo == null || patInfo.SeqInfo.Count == 0
                ? pattern.TestPlanSequences.Count
                : patInfo.SeqInfo.Count;

            #region exists Sequence in patInfo

            for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var sequenceLimits = pattern.MeasPins.Where(p =>
                    p.SequenceIndex == sequenceIndex &&
                    !Regex.IsMatch(p.MeasType, @"Calc|Limit|MeasC", RegexOptions.IgnoreCase)).ToList();
                if (sequenceLimits.Exists(
                    p => p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)))
                    return false;
                if (sequenceLimits.Count > 1)
                    return false;
            }

            #endregion

            return true;
        }

        public static string GetCpuFlag(HardIpReference info, HardIpPattern pattern)
        {
            if (string.IsNullOrEmpty(pattern.Pattern.GetLastPayload()) || pattern.IsNonHardIpBlock) return "false";
            var getCpuFlag = "";

            if (info.MeasSeqStr != "" || info.IsIgnoreComment)
                getCpuFlag += "true";
            else if (pattern.MeasPins.Any(x =>
                (x.MeasType != MeasType.MeasC) & (x.MeasType != MeasType.MeasCalc) &
                (x.MeasType != MeasType.MeasCalcLimit) & (x.MeasType != MeasType.MeasLimit)))
                getCpuFlag += "true";
            else
                getCpuFlag += "false";
            return getCpuFlag;
        }

        public static string GetStoreName(HardIpPattern pattern)
        {
            var storeNameList = new List<string>();
            var pinsInSeq = pattern.MeasPins.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());

            foreach (var pins in pinsInSeq)
            {
                if (pins.Key < 1) continue;
                //concern CP:/FT: with common case
                var storeNameSeqList = pins.Value.Where(p => !p.PinName.Contains("=")).Select(p => p.CusStr).ToList();
                storeNameSeqList.AddRange(pins.Value.Where(p => p.PinName.Contains("CP=")).Select(p => p.CusStr));
                storeNameList.Add(
                    storeNameSeqList.Distinct().Count(p => !string.IsNullOrEmpty(p)) == storeNameSeqList.Count
                        ? string.Join(":", storeNameSeqList)
                        : storeNameSeqList.FirstOrDefault(p => !string.IsNullOrEmpty(p)));
            }

            return storeNameList.Any(p => !string.IsNullOrEmpty(p)) ? string.Join("+", storeNameList) : "";
        }

        public static List<string> DecomposeGroups(string pinGroup)
        {
            var newPinList = new List<string>();
            if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(pinGroup))
                foreach (var pin in HardIpDataMain.TestPlanData.PinGroupList[pinGroup])
                {
                    if (pin.Equals(pinGroup, StringComparison.CurrentCultureIgnoreCase))
                    {
                        newPinList.Add(pin);
                        continue;
                    }

                    newPinList.AddRange(DecomposeGroups(pin));
                }
            else
                newPinList.Add(pinGroup);

            return newPinList;
        }

        public static string GetMeasCPins(HardIpPattern pattern)
        {
            var info = GetHardIpInfo(pattern);
            if (info.CapBitStr != "")
            {
                if (info.CapPinName != "")
                    return info.CapPinName;
                return "JTAG_TDO";
            }

            return "";
        }

        public static bool IsMeasPinInForcePin(string forcePinGroup, string measPinName)
        {
            if (forcePinGroup.Equals(measPinName)) return true;
            if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(forcePinGroup))
            {
                var forcePinGroupList = DecomposeGroups(forcePinGroup);
                if (forcePinGroupList.Contains(measPinName))
                    return true;
                return false;
            }

            return false;
        }

        public static string GetSrcPin(HardIpReference patInfo)
        {
            if (patInfo.SendBit != 0 || patInfo.SendBitStr != "")
            {
                var srcPin = string.IsNullOrEmpty(patInfo.SendPinName) ? "JTAG_TDI" : patInfo.SendPinName;
                return srcPin;
            }

            return "";
        }

        public static string GetPinType(string pinName)
        {
            if (Regex.IsMatch(pinName, "^VDD", RegexOptions.IgnoreCase))
                return "power";
            var name = pinName;
            if (name.Contains("::"))
                name = pinName.Split(':')[0];
            else if (name.Contains(":"))
                name = pinName.Split(':')[1];
            if (HardIpDataMain.TestPlanData == null) return "I/O";
            if (HardIpDataMain.TestPlanData.PinList.ContainsKey(name))
                return HardIpDataMain.TestPlanData.PinList[name];
            return "I/O";
        }

        public static int MeasC_Count(HardIpPattern pattern)
        {
            return pattern.MeasPins.Where(pin => pin.MeasType == "MeasC").Sum(pin => pin.PinCount);
        }

        public static void GetPlanCurrentRange(List<MeasPin> measPins, List<MeasPin> patInfoPins, bool isRepeatLimit)
        {
            if (measPins.Exists(p => p.MeasType.Equals(MeasType.WiSrc) || p.MeasType.Equals(MeasType.WiMeas))) return;

            foreach (var planMeasPin in measPins) planMeasPin.IsUsedPin = false;
            var vocmPins = new List<MeasPin>();

            for (var index = 0; index < patInfoPins.Count; index++)
            {
                var patInfoPin = patInfoPins[index];
                if (patInfoPin.PinName.ToUpper() == "VDD_FIXED_PLL_SOC_ED0_S1")
                {
                }

                var cnt = measPins.Count(x =>
                    ContainsPin(x.PinName, patInfoPin.PinName) &&
                    x.MeasType.ToUpper() == patInfoPin.MeasType.ToUpper() && (x.VisitedTime > 0 || isRepeatLimit) &&
                    !x.IsUsedPin);
                for (var i = 0; i < measPins.Count; i++)
                {
                    var planMeasPin = measPins[i];
                    if (cnt > 0)
                    {
                        if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName) &&
                            patInfoPin.MeasType.Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) &&
                            (planMeasPin.VisitedTime > 0 || isRepeatLimit) && !planMeasPin.IsUsedPin)
                        {
                            patInfoPin.Copy(planMeasPin); //.CurrentRange = planMeasPin.CurrentRange;
                            planMeasPin.IsUsedPin = true;
                            if (planMeasPin.MeasType == "MeasVdiff2")
                                planMeasPin.VisitedTime = planMeasPin.VisitedTime - 2;
                            else
                                planMeasPin.VisitedTime--;
                            break;
                        }
                    }
                    else
                    {
                        if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName) &&
                            patInfoPin.MeasType.Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) &&
                            (planMeasPin.VisitedTime > 0 || isRepeatLimit))
                        {
                            patInfoPin.Copy(planMeasPin); //.CurrentRange = planMeasPin.CurrentRange;
                            planMeasPin.IsUsedPin = true;
                            if (planMeasPin.MeasType == "MeasVdiff2")
                                planMeasPin.VisitedTime = planMeasPin.VisitedTime - 2;
                            else
                                planMeasPin.VisitedTime--;
                            break;
                        }
                    }
                }

                if (patInfoPin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                {
                    var vocmPin = new MeasPin();
                    foreach (var planMeasPin in measPins)
                        if (cnt > 0)
                        {
                            if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName) &&
                                planMeasPin.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                            {
                                vocmPin.Copy(planMeasPin); //.CurrentRange = planMeasPin.CurrentRange;
                                planMeasPin.IsUsedPin = true;
                                vocmPin.PinName = patInfoPin.PinName;
                                vocmPin.SequenceIndex = patInfoPin.SequenceIndex;
                                vocmPins.Add(vocmPin);
                                break;
                            }
                        }
                        else
                        {
                            if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName) &&
                                patInfoPin.MeasType.Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) &&
                                (planMeasPin.VisitedTime > 0 || isRepeatLimit))
                            {
                                vocmPin.Copy(planMeasPin); //.CurrentRange = planMeasPin.CurrentRange;
                                planMeasPin.IsUsedPin = true;
                                vocmPin.SequenceIndex = patInfoPin.SequenceIndex;
                                vocmPin.PinName = patInfoPin.PinName;
                                vocmPins.Add(vocmPin);
                                if (planMeasPin.MeasType == "MeasVdiff2")
                                    planMeasPin.VisitedTime = planMeasPin.VisitedTime - 2;
                                else
                                    planMeasPin.VisitedTime--;
                                break;
                            }
                        }
                }
            }

            patInfoPins.AddRange(vocmPins);
            ResetVisitTime(measPins);
        }

        private static bool ContainsPin(string planPinName, string patInfoPinName)
        {
            if (planPinName == patInfoPinName)
                return true;

            var newPlanPinName = planPinName;
            var newPatInfoPinName = patInfoPinName;
            if (!planPinName.Contains("::") && planPinName.Contains("="))
                newPlanPinName = planPinName.Split('=')[1];
            if (!patInfoPinName.Contains("::") && patInfoPinName.Contains(":"))
                newPatInfoPinName = patInfoPinName.Split(':')[1];
            if (newPlanPinName.Equals(newPatInfoPinName, StringComparison.OrdinalIgnoreCase))
                return true;
            //if (newPlanPinName == newPatInfoPinName)
            //    return true;

            if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(planPinName))
            {
                var pinGroupList = DecomposeGroups(planPinName);
                if (pinGroupList.Contains(newPatInfoPinName))
                    return true;
                return false;
            }

            return false;
        }

        private static void ResetVisitTime(List<MeasPin> measPins)
        {
            foreach (var planMeasPin in measPins) planMeasPin.VisitedTime = planMeasPin.PinCount;
        }

        public static string GetCusStrDigCapData(HardIpPattern pattern)
        {
            #region Get DsscOut from pattern info && check if Disable_MeasC_Split

            var info = GetHardIpInfo(pattern);

            if (Regex.IsMatch(info.DsscOut, "Reg_assign:", RegexOptions.IgnoreCase)) return info.DsscOut;

            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.KeepDsscOut, RegexOptions.IgnoreCase) &&
                info.CapBitStr != "" && info.CapBitName != "" && info.DsscOut != "")
            {
                var bitStrArray = info.CapBitStr.Split('+');
                var bitNameArray = info.CapBitName.Split('+');
                var oriDsscOut = Regex.Replace(info.DsscOut, @"DSSC_OUT.\s*", "", RegexOptions.IgnoreCase).Trim(',');
                var dsscOutArray = oriDsscOut.Split(',');
                var newDsscOut = new List<string>();
                var index = 0;
                var nameIndex = 0;
                foreach (var bitStr in bitStrArray)
                {
                    var width = Convert.ToInt32(bitStr.Split('_')[1]);
                    var dsscOutWidth = Convert.ToInt32(dsscOutArray[index].Split(':')[0]);
                    if (width == dsscOutWidth)
                    {
                        newDsscOut.Add(width + ":" + bitNameArray[nameIndex]);
                    }
                    else
                    {
                        while (width > dsscOutWidth)
                        {
                            index++;
                            var nextDsscOutWidth = Convert.ToInt32(dsscOutArray[index].Split(':')[0]);
                            dsscOutWidth += nextDsscOutWidth;
                        }

                        if (width == dsscOutWidth)
                            newDsscOut.Add(width + ":" + bitNameArray[nameIndex]);
                    }

                    index++;
                    nameIndex++;
                }

                info.DsscOut = "DSSC_OUT," + string.Join(",", newDsscOut);
            }

            var dsscOut = Regex.Replace(info.DsscOut.Trim(','), "DSSC_OUT:", "DSSC_OUT,", RegexOptions.IgnoreCase)
                .Split(',').ToList();

            #endregion

            #region When set "ignore_patt_comment", user can change bits of measC by testPlan

            //bool ignorePatInfo = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase);
            //bool ignorePatMeasC = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatMeasC, RegexOptions.IgnoreCase);
            //PinSeqChecker capBits = new PinSeqChecker();
            //if (ignorePatMeasC)
            //{
            //    if (capBits.CapBitsInTp(pattern) == info.CapBit && info.CapBit != 0)
            //    {
            //        dsscOut = new List<string> { "DSSC_OUT" };
            //        foreach (
            //            var pinMeasC in
            //                pattern.MeasPins.Where(
            //                    p => p.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase)))
            //        {
            //            dsscOut.Add(string.Format("{0}:{1}", pinMeasC.CapBit, pinMeasC.CusStr));
            //        }
            //    }
            //}
            //else if (ignorePatInfo)
            //{
            //    if (capBits.CapBitsInTp(pattern) == info.CapBit && info.CapBit != 0)
            //    //Total bits in TestPlan and info are the same
            //    {
            //        bool isCapBitsMatch = IsCapBitsMatch(pattern, info);
            //        if ((dsscOut.Count - 1) == MeasC_Count(pattern) && MeasC_Count(pattern) != 0 &&
            //            !isCapBitsMatch)
            //        {
            //            dsscOut.RemoveRange(0, dsscOut.Count);
            //            dsscOut.Add("DSSC_OUT");
            //            foreach (var measCPin in pattern.MeasPins.FindAll(s => s.MeasType == MeasType.MeasC).ToList())
            //            {
            //                dsscOut.Add(measCPin.CapBit + ":MeasC_" + (dsscOut.Count));
            //            }
            //        }
            //    }
            //}

            #endregion

            #region Replace DSSC OUT by testName/CusStr in testPlan

            var capPinIndex = 1;
            foreach (var pin in pattern.MeasPins)
            {
                if (pin.MeasType != MeasType.MeasC)
                    continue;

                if (dsscOut.Count > capPinIndex)
                {
                    var strArr = dsscOut[capPinIndex].Split(':');
                    if (strArr.Length == 2)
                    {
                        if (!string.IsNullOrEmpty(pin.TestName))
                            strArr[1] = pin.TestName;
                        dsscOut[capPinIndex] = string.Join(":", strArr.ToList());
                        if (!string.IsNullOrEmpty(pin.CusStr))
                            dsscOut[capPinIndex] = dsscOut[capPinIndex] + ":" + pin.CusStr;
                    }
                }

                capPinIndex++;
            }

            var capData = string.Join(",", dsscOut);

            #endregion

            var trimOrCapName = GetDigCapNameByMiscInfo(pattern.MiscInfo);
            var allCusStr = pattern.MeasPins
                .Where(pin => pin.MeasType == MeasType.MeasC && !string.IsNullOrEmpty(pin.CusStr))
                .Select(cPin => cPin.CusStr).ToList();
            if (trimOrCapName != "")
                capData = (trimOrCapName + "&" + capData).Trim('&');
            else if (allCusStr.Distinct().Count() == 1 && !pattern.MeasPins.Any(pin =>
                pin.MeasType == MeasType.MeasC && string.IsNullOrEmpty(pin.CusStr)))
                capData = (allCusStr[0] + "&" + capData).Trim('&');

            return capData == "" ? "" : capData;
        }

        public static string GetFlowForLoopIntegerName(HardIpPattern pattern)
        {
            var flowForLoopIntegerName = "";

            if (pattern.SweepCodes.Count > 0)
            {
                var sweepList = new List<string>();
                foreach (var sweepItems in pattern.SweepCodes)
                {
                    flowForLoopIntegerName = string.Format("SrcCodeIndx{0};", sweepItems.Key);
                    foreach (var sweepCode in sweepItems.Value)
                        if (string.IsNullOrEmpty(sweepCode.Misc))
                            flowForLoopIntegerName += sweepCode.SendBitName + ":" + sweepCode.Width + ":" +
                                                      sweepCode.Start + ":" + sweepCode.Step + ";";
                        else
                            flowForLoopIntegerName += sweepCode.SendBitName + ":" + sweepCode.Width + ":" +
                                                      sweepCode.Start + ":" + sweepCode.Misc + ";";
                    sweepList.Add(flowForLoopIntegerName.Trim(';'));
                }

                return string.Join("|", sweepList);
            }

            return flowForLoopIntegerName.Trim(';');
        }

        public static string GetDigCapNameByMiscInfo(string miscInfo)
        {
            var digCapNameReg = HardIpConstData.DigCapName + @":(?<name>.*)";
            foreach (var item in miscInfo.Split(';'))
                if (Regex.IsMatch(item, digCapNameReg, RegexOptions.IgnoreCase))
                    return Regex.Match(item, digCapNameReg, RegexOptions.IgnoreCase).Groups["name"].ToString();
            return "";
        }

        public static string GetVbtNameByPattern(HardIpPattern pattern)
        {
            if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget)) pattern.VbtTypes.Add(PlanType.Trim);
            if (pattern.MeasPins.Exists(p =>
                p.MeasType.Equals(MeasType.MeasWait, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.WiSrc, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.WiMeas, StringComparison.OrdinalIgnoreCase)))
                pattern.VbtTypes.Add(PlanType.Rf);

            foreach (var meas in pattern.MeasPins)
            {
                var bbVbtName = GetBbFunc(meas.RfInstrumentSetup);
                if (!string.IsNullOrEmpty(bbVbtName))
                    return bbVbtName;
            }


            var isCalcArg = Regex.IsMatch(pattern.MiscInfo, "CalcArg:", RegexOptions.IgnoreCase);
            foreach (var info in pattern.MiscInfo.Split(';'))
            {
                if (!Regex.IsMatch(info.Trim(), "^(" + HardIpConstData.Vbt + "|calc" + @")\s*\:\s*\w+",
                    RegexOptions.IgnoreCase))
                    continue;

                var mappingPair = HardIpDataMain.ConfigData.VbtNameMapping.FirstOrDefault(a =>
                    a.Key.Equals(info.Trim(), StringComparison.OrdinalIgnoreCase));
                if (mappingPair.Value != null)
                    return mappingPair.Value;


                if (!isCalcArg)
                {
                    if (!string.IsNullOrEmpty(TestProgram.VbtFunctionLib.GetFunctionByName(info.Split(':')[1])
                        .Parameters))
                        return info.Split(':')[1];

                    if (Regex.IsMatch(info, "^calc:|^Meas:", RegexOptions.IgnoreCase))
                        return Regex.Replace(info, "^calc:|^Meas:", "", RegexOptions.IgnoreCase);
                }
            }

            if (pattern.SheetName.StartsWith("PrefixLCD", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
                    return VbtFunctionLib.DvdcTrim;
                return VbtFunctionLib.LcdMeas;
            }

            if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget)) return VbtFunctionLib.DvdcTrim;
            if (pattern.MeasPins.Exists(p =>
                p.MeasType.Equals(MeasType.WiSrc, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.WiMeas, StringComparison.OrdinalIgnoreCase)))
            {
                if (ProjectConfigSingleton.Instance().GetProjectConfigValue("Wireless", "RFItem") == "UniqueVBT")
                {
                    var vbtName = string.IsNullOrEmpty(pattern.WirelessData.TrimTarget) ? "RFFunc" : "RFTrim";
                    var setupInfos = GetInstrumentInfo(pattern, "RFInstSetup").Split('|');
                    var instList = new List<string>();
                    foreach (var setupInfo in setupInfos)
                    {
                        var subInfos = new Dictionary<string, int>();
                        var instCount = new List<string>();
                        foreach (var subInfo in setupInfo.Split('+'))
                        {
                            if (!subInfo.Contains("#")) continue;
                            var instKey = subInfo.Split('#')[1].Split('_')[0];
                            if (!subInfos.ContainsKey(instKey))
                                subInfos.Add(instKey, 0);
                            subInfos[instKey]++;
                        }

                        foreach (var inst in subInfos) instCount.Add(string.Format("{0}x{1}", inst.Key, inst.Value));
                        instList.Add(string.Join("", instCount));
                    }

                    vbtName = string.Format("{0}_{1}", vbtName, string.Join("_", instList));
                    if (!VbtFunctionLib.GeneratedVbtFunctionDic.ContainsKey(vbtName))
                        VbtFunctionLib.GeneratedVbtFunctionDic.Add(vbtName, 0);

                    pattern.CustomVbName = vbtName;
                }

                return VbtFunctionLib.RfFunc;
            }


            if (pattern.MeasPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff2)))
                return VbtFunctionLib.VdiffFunc;

            return "";
        }

        private static string GetBbFunc(string forceStr)
        {
            const string bbFunc = "BBFunc";
            foreach (var force in forceStr.Split(';'))
                if (Regex.IsMatch(force, bbFunc, RegexOptions.IgnoreCase))
                    return force.Split('=')[1];
            return "";
        }

        public static string GetSpecifyInfo(string dcInfo, string key)
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

        public static string GetAcSelector(string labelVoltage, string acInfo)
        {
            var acSelector = "";
            foreach (var info in acInfo.Split(';'))
            {
                var acArray = info.Split(':');
                if (acArray.Length != 3) continue;
                if (acArray[1].Equals(labelVoltage, StringComparison.OrdinalIgnoreCase)) acSelector = acArray[2];
                if (acArray[1].Equals(HardIpConstData.LabelAll, StringComparison.OrdinalIgnoreCase))
                    acSelector = acArray[2];
            }

            if (acSelector.Equals(HardIpConstData.SelectMax, StringComparison.OrdinalIgnoreCase))
                acSelector = HardIpConstData.SelectMax;
            if (acSelector.Equals(HardIpConstData.SelectMin, StringComparison.OrdinalIgnoreCase))
                acSelector = HardIpConstData.SelectMin;
            if (acSelector.Equals(HardIpConstData.SelectTyp, StringComparison.OrdinalIgnoreCase))
                acSelector = HardIpConstData.SelectTyp;

            return acSelector;
        }

        public static string GetDcSelector(string labelVoltage, string dcInfo)
        {
            var dcSelector = "";
            foreach (var info in dcInfo.Split(';'))
            {
                var dcArray = info.Split(':');
                if (dcArray.Length != 3) continue;
                if (dcArray[1].Equals(labelVoltage, StringComparison.OrdinalIgnoreCase)) dcSelector = dcArray[2];
                if (dcArray[1].Equals(HardIpConstData.LabelAll, StringComparison.OrdinalIgnoreCase))
                    dcSelector = dcArray[2];
            }

            if (dcSelector.Equals(HardIpConstData.SelectMax, StringComparison.OrdinalIgnoreCase))
                dcSelector = HardIpConstData.SelectMax;
            if (dcSelector.Equals(HardIpConstData.SelectMin, StringComparison.OrdinalIgnoreCase))
                dcSelector = HardIpConstData.SelectMin;
            if (dcSelector.Equals(HardIpConstData.SelectTyp, StringComparison.OrdinalIgnoreCase))
                dcSelector = HardIpConstData.SelectTyp;

            return dcSelector;
        }

        public static List<Timing> GetTimingsByAc(string acInfo)
        {
            var timings = new List<Timing>();
            foreach (var info in acInfo.Split(';'))
            {
                var acArray = info.Split(':');
                if (!info.Contains("AC") && !(acArray.Length == 3 && acArray.Length == 4))
                    continue;
                var timing = new Timing();
                timing.Name = acArray[1];
                if (acArray.Length == 3)
                    timing.SuffixAcSpecName = acArray[2];
                else if (acArray.Length == 4)
                    timing.SuffixAcSpecName = acArray[3];

                // var pin = NwireSingleton.Instance().SettingInfo.NwirePins.Find(s => s.OutClk.Equals(timing.Name, StringComparison.OrdinalIgnoreCase));
                // if (pin != null) timing.Name = pin.CreatePinNameWithDiff();

                var acValue = acArray[2];

                if (acValue.Contains(","))
                {
                    var data = acValue.Split(',');
                    var value = GetFreq(data[0]);
                    var category = data[1];
                    if (category.Equals("Typ", StringComparison.OrdinalIgnoreCase))
                        timing.Type = value;
                    else if (category.Equals("Min", StringComparison.OrdinalIgnoreCase))
                        timing.Min = value;
                    else if (category.Equals("Max", StringComparison.OrdinalIgnoreCase))
                        timing.Max = value;
                }
                else
                {
                    var value = GetFreq(acValue);
                    timing.Type = value;
                    timing.Min = value;
                    timing.Max = value;
                }

                timings.Add(timing);
            }

            return timings;
        }

        public static string GetInstNameSubStr(string miscInfo)
        {
            var instNameSubStr = "";
            foreach (var info in miscInfo.Split(';'))
                if (info.ToLower().Contains(HardIpConstData.InstNameSubStr.ToLower()) && info.Contains(":"))
                {
                    var misArr = info.Split(':');
                    if (misArr.Length == 2 &&
                        misArr[0].Equals(HardIpConstData.InstNameSubStr, StringComparison.OrdinalIgnoreCase))
                    {
                        instNameSubStr = misArr[1];
                        break;
                    }
                }

            return instNameSubStr;
        }

        public static string GetEnvFromPattern(HardIpPattern pattern, bool isBinTableUse = false)
        {
            if (isBinTableUse)
            {
                if (pattern.NoBinOutStr.Equals("All", StringComparison.OrdinalIgnoreCase))
                    return "NoBinOut";
                if (HardIpConstData.LabelVolList.Exists(p =>
                    Regex.IsMatch(pattern.NoBinOutStr, p, RegexOptions.IgnoreCase)))
                    return "SpecialBinOut";
            }

            var patternName = pattern.Pattern.GetLastPayload();
            if (patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(pattern.Pattern.GetLastPayload(), HardIpConstData.RegInsInPattern,
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

        //public static bool GetStatusInPatternList(List<PatternData> patternList, string patternName, out string status)
        //{
        //    status = "";
        //    if (!patternList.ContainsKey(patternName))
        //        status = "MissPattInPattList";
        //    else if (patternList[patternName].TimeSetVersion.ToLower() == "na")
        //        status = "MissTimesetInPattList";
        //    else if (patternList[patternName].FileVersion == "NA")
        //        status = "MissFileVersionInPattList";
        //    if (status != "")
        //        return false;
        //    return true;
        //}

        //public static string GetTtrEnable(string ttrStr, string voltage)
        //{
        //    ttrStr = ttrStr.Replace(" ", "");
        //    var allJobs = HardIpDataMain.TestPlanData.AllJobs.ToList();
        //    foreach (var ttr in ttrStr.Split(';'))
        //    {
        //        if (ttr == "")
        //            continue;
        //        if (!string.IsNullOrEmpty(voltage) && ttr.Contains(voltage))
        //        {
        //            // NV/LV/HV
        //            if (!ttr.Contains(":"))
        //                return "";
        //            // [JobName]:[NV/LV/HV]
        //            var jobName = ttr.Split(':')[0];
        //            allJobs = allJobs.Where(a => !Regex.IsMatch(a, jobName)).ToList();
        //        }
        //        else
        //        {
        //            // eg: CP1 means that NV,HV,LV should be removed from CP1
        //            allJobs = allJobs.Where(a => !Regex.IsMatch(a, ttr)).ToList();
        //        }
        //    }

        //    return string.Join(",", allJobs);
        //}

        public static Dictionary<string, string> GetRelaySetting(HardIpPattern pattern,
            Dictionary<string, string> exSetting)
        {
            var newRelaySetting = new Dictionary<string, string>();
            var actualRelaySetting = new Dictionary<string, string>();
            var settingInMisc = GetsRelaysInMisc(pattern.MiscInfo);
            var relayGroups = from item in settingInMisc group item by item.Job into g select g;
            foreach (var relayGroup in relayGroups)
            {
                var job = relayGroup.Key;
                var relayOnSetting = string.Join("_",
                    relayGroup.ToList().Where(s => s.Status == RelayStatus.On).Select(s => s.Name.Replace(",", "_"))
                        .ToList());
                var relayOffSetting = string.Join("_",
                    relayGroup.ToList().Where(s => s.Status == RelayStatus.Off).Select(s => s.Name.Replace(",", "_"))
                        .ToList());
                if (!string.IsNullOrEmpty(relayOnSetting))
                    relayOnSetting = "RelayOn_" + relayOnSetting;
                if (!string.IsNullOrEmpty(relayOffSetting))
                    relayOffSetting = "RelayOff_" + relayOffSetting;
                newRelaySetting.Add(job, (relayOnSetting + "_" + relayOffSetting).Trim('_'));
            }

            pattern.NewRelaySetting = newRelaySetting;
            var jobs = newRelaySetting.Keys.ToList();
            jobs.AddRange(exSetting.Keys);
            jobs = jobs.Distinct().ToList();
            foreach (var job in jobs)
            {
                string exSettingInJob;
                string newSettingInJob;
                exSetting.TryGetValue(job, out exSettingInJob);
                newRelaySetting.TryGetValue(job, out newSettingInJob);
                if (string.IsNullOrEmpty(exSettingInJob))
                {
                    actualRelaySetting.Add(job, newSettingInJob);
                    continue;
                }

                if (exSettingInJob == newSettingInJob) continue;
                actualRelaySetting.Add(job, (ReverseRelaySetting(exSettingInJob) + ";" + newSettingInJob).Trim(';'));
            }

            pattern.RelaySetting = actualRelaySetting;
            pattern.NewRelaySetting = newRelaySetting;
            return newRelaySetting;
        }

        public static List<HardIpRelay> GetsRelaysInMisc(string miscInfo)
        {
            var settingInMisc = new List<HardIpRelay>();
            foreach (var param in miscInfo.Split(';'))
            {
                var text = param.Trim();
                if (!text.Contains(":") ||
                    !(text.StartsWith(HardIpConstData.RelayOn, StringComparison.CurrentCultureIgnoreCase) ||
                      text.StartsWith(HardIpConstData.RelayOff, StringComparison.CurrentCultureIgnoreCase)))
                    continue;
                var relayArr = param.Split(':');
                var status = relayArr[0].ToLower() == "relayon" ? RelayStatus.On : RelayStatus.Off;
                var job = param.Split(':').Length == 3 ? param.Split(':')[2].ToUpper() : "";
                var relay = new HardIpRelay(job, relayArr[1], status);
                settingInMisc.Add(relay);
            }

            return settingInMisc;
        }

        public static string ReverseRelaySetting(string setting)
        {
            var relayOn = Regex.Match(setting, @"(?<relayon>RelayOff[^R]+).*").Groups["relayon"].ToString()
                .Replace("Off", "On");
            var relayOff = Regex.Match(setting, @"(?<relayoff>RelayOn[^R]+).*").Groups["relayoff"].ToString()
                .Replace("On", "Off");
            return (relayOn + "_" + relayOff).Trim('_');
        }

        public static bool IsForceType(HardIpPattern pattern, string type)
        {
            return pattern.MeasPins.Any(x =>
                x.ForceConditions.Any(y => y.ForcePins.Any(z => z.ForceType.StartsWith(type))));
        }

        public static bool IsMeasE(HardIpPattern pattern)
        {
            return pattern.MeasPins.Any(pin => pin.MeasType.Equals("measE", StringComparison.CurrentCultureIgnoreCase));
        }

        public static int MeasECount(HardIpPattern pattern)
        {
            return pattern.MeasPins
                       .FindAll(pin => pin.MeasType.Equals("measE", StringComparison.CurrentCultureIgnoreCase)).Count /
                   2;
        }

        public static bool IsMeasVdiff2(HardIpPattern pattern)
        {
            var patInfo = GetHardIpInfo(pattern);
            return pattern.MeasPins.Any(pin => pin.MeasType.ToLower() == "measvdiff2") ||
                   patInfo != null && patInfo.MeasVdiff2PinList.Count > 0;
        }

        public static bool IsMeasVdiff(HardIpPattern pattern)
        {
            var patInfo = GetHardIpInfo(pattern);
            return pattern.MeasPins.Any(pin =>
                       pin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)) ||
                   patInfo != null && patInfo.MeasVdiffPinList.Count > 0;
        }

        public static bool IsMeasI2(HardIpPattern pattern)
        {
            return pattern.MeasPins.Any(pin => pin.MeasType.Equals(MeasType.MeasI2));
        }

        public static bool IsRepeatLimit(string miscInfo)
        {
            return miscInfo.Split(';').ToList()
                .Exists(s => Regex.IsMatch(s, HardIpConstData.RepeatLimit, RegexOptions.IgnoreCase));
        }

        public static int GetRepeatLimitCount(string miscInfo)
        {
            var loop = 0;
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 &&
                    Regex.IsMatch(assignArr[0], HardIpConstData.RepeatLimit, RegexOptions.IgnoreCase))
                    int.TryParse(assignArr[1].Trim(), out loop);
            }

            return loop;
        }

        public static List<string> GetOpCode(HardIpPattern pattern, string opType)
        {
            var miscInfo = pattern.MiscInfo;
            var opCodeListA = new List<string>(); //After the pattern
            var opCodeListB = new List<string>(); //Before the pattern
            foreach (var info in miscInfo.Split(';'))
            {
                if (!Regex.IsMatch(info, @"opcode\s*\:\w+\s*(\:.*)?", RegexOptions.IgnoreCase))
                    continue;
                var code = Regex.Match(info, @"opcode\s*\:(?<code>(\w+))\s*(\:.*)?", RegexOptions.IgnoreCase)
                    .Groups["code"].ToString();
                var name = Regex.Match(info, @"opcode\s*\:\w+\s*(\:\s*(?<name>(.*)))?", RegexOptions.IgnoreCase)
                    .Groups["name"].ToString();
                if (opType == "B" && !(code.ToLower() == "endif" || code.ToLower() == "next") &&
                    pattern.DivideIndex != 2 && pattern.FlowControlFlag != 1)
                    opCodeListB.Add(code + ":" + name);
                if (opType == "A" && (code.ToLower() == "endif" || code.ToLower() == "next") &&
                    pattern.DivideIndex != 1 && pattern.FlowControlFlag != 0)
                    opCodeListA.Add(code + ":" + name);
            }

            if (opType == "A")
                return opCodeListA;
            return opCodeListB;
        }

        public static string GetFlagNoBinStr(string miscInfo, string voltage)
        {
            foreach (var item in miscInfo.Split(';'))
                if (Regex.IsMatch(item, HardIpConstData.NoBin, RegexOptions.IgnoreCase))
                {
                    var noBinVoltage = item.Split(':').ToList();
                    if (noBinVoltage.Count > 1 && Regex.IsMatch(noBinVoltage[1], voltage))
                        return "No";
                }

            //string noBinVoltages = Regex.Match(miscInfo, HardIpConstData.NoBin + @":(?<setting>)", RegexOptions.IgnoreCase).Groups["setting"].ToString();
            //if (Regex.IsMatch(noBinVoltages, voltage, RegexOptions.IgnoreCase))
            //return "No";
            return "";
        }

        public static string GetFlagNoBinUseLimitStr(string miscInfo, string voltage)
        {
            foreach (var item in miscInfo.Split(';'))
                if (Regex.IsMatch(item, HardIpConstData.NoBinUseLimit + "|" + HardIpConstData.NoBin,
                    RegexOptions.IgnoreCase))
                {
                    var noBinVoltage = item.Split(':').ToList();
                    if (noBinVoltage.Count > 1 && Regex.IsMatch(noBinVoltage[1], voltage))
                        return "No";
                }

            //string noBinVoltages = Regex.Match(miscInfo, HardIpConstData.NoBinUseLimit + "|" + HardIpConstData.NoBin + @":(?<setting>)", RegexOptions.IgnoreCase).Groups["setting"].ToString();
            //if (Regex.IsMatch(noBinVoltages, voltage, RegexOptions.IgnoreCase))
            //    return "No";
            return "";
        }

        public static void InitialMeasC(HardIpPattern pattern, HardIpReference patInfo)
        {
            if (patInfo.DsscOut != "")
            {
                var planCount = MeasC_Count(pattern);
                var infoCount = patInfo.DsscOut.Trim(',').Split(',').Length - 1;
                if (infoCount > planCount && !Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatMeasC,
                    RegexOptions.IgnoreCase))
                {
                    var pinName = patInfo.CapPinName == "" ? "JTAGTDO" : patInfo.CapPinName;
                    for (var i = 0; i < infoCount - planCount; i++)
                    {
                        var pin = new MeasPin(pinName, "MeasC");
                        pin.PinCount = 1;
                        pattern.MeasPins.Add(pin);
                    }
                }
            }
        }

        public static BinNumberRuleRow GetHardIpBin(HardIpPattern pattern)
        {
            const string index5 = "HardIP_others";
            BinNumberRuleRow binRange;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.HardIp, index5);

            para.Condition = Regex.Replace(pattern.SheetName, "HardIP_|Wireless_|LCD_", "", RegexOptions.IgnoreCase)
                .Replace("_", "");
            var found = BinNumberSingleton.Instance().GetBinNumDefRow(para, out binRange);
            if (found)
                return binRange;
            para.Condition = index5;
            found = BinNumberSingleton.Instance().GetBinNumDefRow(para, out binRange);
            if (found)
                return binRange;
            const string errorMessage = "Missing bin number setting";
            EpplusErrorManager.AddError(HardIpErrorType.MissingBinNum, ErrorLevel.Error, pattern.SheetName,
                pattern.RowNum, errorMessage, para.Condition);
            return binRange;
        }

        public static string GetFreq(string srcFreq)
        {
            var result = srcFreq;
            const string lStrValuePattern = @"^(?<str>\d*[.]?\d*)[a-zA-Z]+$";
            const string lStrUnitPattern = @"^\d*[.]?\d*(?<str>[a-zA-Z]+)$";
            if (Regex.IsMatch(result, lStrUnitPattern))
            {
                var lStrUnit = Regex.Match(result, lStrUnitPattern).Groups["str"].ToString().Trim().ToUpper();
                var lStrValue = Regex.Match(result, lStrValuePattern).Groups["str"].ToString().Trim();
                TryToConvertToHz(lStrValue, lStrUnit, out result);
            }

            return result;
        }

        public static bool TryToConvertToHz(string value, string unit, out string outputValue)
        {
            double lDValue;
            outputValue = string.Empty;
            if (double.TryParse(value, out lDValue) == false) return false;
            if (unit.Equals(CommonConst.UnitHz, StringComparison.OrdinalIgnoreCase))
                outputValue = value;
            else if (unit.Equals(CommonConst.UnitKhz, StringComparison.OrdinalIgnoreCase))
                outputValue = (lDValue * 1e3).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(CommonConst.UnitMhz, StringComparison.OrdinalIgnoreCase))
                outputValue = (lDValue * 1e6).ToString(CultureInfo.InvariantCulture);
            return true;
        }

        public static string GetCalculation(string miscInfo, string testName = "")
        {
            var calculation = "";
            var calculationList = new List<string>();
            if (miscInfo.IndexOf(HardIpConstData.Calc + ":", StringComparison.OrdinalIgnoreCase) != -1)
            {
                var algorithm = "";
                var paras = "";
                var flag = false;
                var setFlag = false;
                foreach (var item in miscInfo.Split(';'))
                {
                    if (item.StartsWith(HardIpConstData.Calc + ":") && !flag)
                    {
                        algorithm = item.Split(':')[1];
                        flag = true;
                    }

                    if (item.StartsWith(HardIpConstData.CalcParameter + ":") && flag)
                    {
                        paras = item.Split(':')[1];
                        flag = false;
                        setFlag = true;
                    }

                    if (setFlag && algorithm != "" && paras != "")
                    {
                        calculationList.Add("Alg:" + testName + ":" + algorithm + "(" + paras + ")");
                        setFlag = false;
                    }
                }
            }

            if (calculationList.Any()) calculation = string.Join(";", calculationList);
            return calculation;
        }

        public static string TrimSpace(string input)
        {
            return input.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Trim();
        }

        public static List<string> GetMeasStrByPlan(HardIpPattern pattern)
        {
            var measList = new List<string>();
            for (var seqIndex = 1; seqIndex <= pattern.TestPlanSequences.Count; seqIndex++)
            {
                var firstOrDefault = pattern.MeasPins.FirstOrDefault(s =>
                    s.SequenceIndex == seqIndex && s.MeasType != "MeasC" && s.PinName != "FakePin" && s.MeasType != "");
                if (firstOrDefault != null)
                {
                    if (firstOrDefault.MeasType.Equals(MeasType.WiMeas))
                        measList.Add("A");
                    else if (firstOrDefault.MeasType.Equals(MeasType.WiSrc))
                        measList.Add("G");
                    else if (firstOrDefault.MeasType.Equals(MeasType.MeasWait))
                        measList.Add("W");
                    else
                        measList.Add(firstOrDefault.MeasType.Replace("Meas", ""));
                }

                else if (pattern.TestPlanSequences[seqIndex - 1].ForceCondition.Count > 0)
                {
                    measList.Add("N");
                }
            }

            if (measList.Count == 1 && measList[0] == "N")
                measList.Clear();
            return measList;
        }

        public static string GenDiffGroupName(string diffPinName, bool isNeedGenPinGroup)
        {
            string pPin;
            string nPin;
            var groupName = "";
            if (!diffPinName.Contains("::")) return diffPinName;
            if (!isNeedGenPinGroup) return string.Join(",", Regex.Split(diffPinName, "::"));
            if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
            {
                var pair = diffPinName.Split(new[] {"::"}, StringSplitOptions.None);
                groupName = TestProgram.IgxlWorkBk.PinMapPair.Value.GetDiffGroupName(pair);
            }

            if (groupName == "")
                DiffPinPosAndNeg(diffPinName, out pPin, out nPin, out groupName);
            if (groupName == "")
                return diffPinName;
            return groupName;
        }

        public static bool DiffPinPosAndNeg(string diffPins, out string pos, out string neg, out string groupName)
        {
            pos = "";
            neg = "";
            groupName = "";
            if (!diffPins.Contains("::"))
                return false;

            var pair = diffPins.Split(new[] {"::"}, StringSplitOptions.None);
            //   DiffPairConfig config = XmlSer<DiffPairConfig>.LoadXml(Directory.GetCurrentDirectory() + "/Config/DiffPairConfig.xml");
            for (var i = 0; i < pair.Length; i++)
                pair[i] = pair[i].Trim();

            groupName = "";

            return false;
        }

        public static string GetDigDataWidth(string sendBitStr, string defaultValue = "")
        {
            var reg = @"[a-zA-Z]+\d+_(?<num>(\d+))";
            var matches = Regex.Matches(sendBitStr, reg, RegexOptions.IgnoreCase);
            if (matches.Count == 0)
                return string.Empty;
            var width = matches[0].Groups["num"].ToString();
            if (matches.Cast<Match>().Any(match => !match.Groups["num"].ToString().Equals(width))) return defaultValue;
            return width;
        }

        public static string GetPpmuPin(HardIpPattern pattern, HardIpReference info)
        {
            var measPPmu = "";
            var seqMeas = new List<string>();
            var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount <= 0) return measPPmu;

            for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var measPinList = pattern.MeasPins.Where(a =>
                    a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC" && a.MeasType != MeasType.MeasLimit &&
                    a.MeasType != "").ToList();
                var internalMeas = measPinList.Select(measPin => measPin.PinName).ToList();
                seqMeas.Add(string.Join(",", internalMeas));
            }

            measPPmu = string.Join("+", seqMeas)
                .Replace("::", ","); //Convert P1_P::P1_N to P1_P,P1_N(differential pins)

            return measPPmu;
        }

        public static string GetForceV(HardIpPattern pattern, HardIpReference info)
        {
            if (IsForceType(pattern, "V"))
            {
                var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                var seqList = new List<string>();
                if (seqCount > 0)
                {
                    for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var measPinList = pattern.MeasPins
                            .Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != MeasType.MeasC).ToList();
                        var isAllMeasR2 = IsAllTheSameType(measPinList, MeasType.MeasR2);
                        var isAllMeasI = IsAllTheSameType(measPinList, MeasType.MeasI);
                        var isAllMeasV = IsAllTheSameType(measPinList, MeasType.MeasV);
                        var forceDelimiter = isAllMeasI || isAllMeasV ? ":" : ",";

                        #region measR2

                        if (isAllMeasR2)
                        {
                            var fieldList = new List<string>();
                            foreach (var measPin in measPinList)
                            {
                                var forceVPerPin = new List<string>();
                                if (measPin.ForceConditions.Count > 0)
                                {
                                    foreach (var condition in measPin.ForceConditions)
                                    foreach (var forcePin in condition.ForcePins)
                                        if (forcePin.ForceType == "V")
                                            if (forcePin.PinName == measPin.PinName)
                                                forceVPerPin.Add(DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                                }
                                else
                                {
                                    forceVPerPin.Add("");
                                }

                                bool errorFlag;
                                fieldList.Add(GetByPair(forceVPerPin, "&", ":", out errorFlag));
                                if (errorFlag)
                                {
                                    var forceIndex =
                                        HardIpDataMain.TestPlanData.PlanHeaderIdx[pattern.SheetName]["forceIndex"];
                                    var errorMessage = "The count of MeasR2 Pin = " + measPin.PinName +
                                                       " force condition must be even !!!";
                                    if (!EpplusErrorManager.GetErrors().Any(x =>
                                        x.SheetName == pattern.SheetName && x.RowNum == measPin.RowNum &&
                                        x.ColNum == forceIndex && x.Message == errorMessage))
                                        EpplusErrorManager.AddError(HardIpErrorType.WrongForceCondition,
                                            ErrorLevel.Error, pattern.SheetName, measPin.RowNum, forceIndex,
                                            errorMessage);
                                }
                            }

                            seqList.Add(string.Join(",", fieldList));
                        }

                        #endregion

                        #region others

                        else
                        {
                            var fieldList = new List<string>();
                            foreach (var measPin in measPinList)
                            {
                                var forceVPerPin = new List<string>();
                                if (measPin.ForceConditions.Count > 0)
                                    foreach (var condition in measPin.ForceConditions)
                                    foreach (var forcePin in condition.ForcePins)
                                    {
                                        var measPinName = measPin.PinName.Split(':').Length == 2
                                            ? measPin.PinName.Split(':')[1]
                                            : measPin.PinName;
                                        if (!IsMeasPinInForcePin(forcePin.PinName, measPinName) ||
                                            forcePin.Type == ForceConditionType.Others)
                                            continue;
                                        //if (forcePin.ForceType == "V")
                                        //{
                                        if (forcePin.ForceJob == "")
                                            forceVPerPin.Add(DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                                        else
                                            forceVPerPin.Add(forcePin.ForceJob + ":" +
                                                             DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                                        //}
                                    }
                                else
                                    forceVPerPin.Add("");

                                fieldList.Add(string.Join(forceDelimiter, forceVPerPin));
                            }

                            seqList.Add(string.Join(",", fieldList));
                        }

                        #endregion
                    }

                    var forceV = string.Join("|", seqList);

                    //if (HardIpDataMain.ConfigData.SeqDelimiter == "|")
                    //    forceV = forceV.Replace("&", "|");
                    //else
                    //    if (!forceV.Contains("+")) forceV = forceV.Replace("&", "+");

                    return DataConvertor.RemoveDummyForceV(forceV, @",|\+|\&|\|");
                }

                return "";
            }

            return "";
        }

        public static string GetForceI(HardIpPattern pattern, HardIpReference info)
        {
            if (IsForceType(pattern, "I"))
            {
                var forceI = "";
                var isMeasE = IsMeasE(pattern);
                var isAllMEasI = IsAllMeasI(pattern);
                var forceDelimiter = isAllMEasI ? ':' : ',';
                var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                if (seqCount > 0)
                {
                    var forceIe1 = ""; //First condition for MeasE
                    var forceIe2 = ""; //Second condition for MeasE
                    for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var measPinList = pattern.MeasPins
                            .Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC").ToList();
                        foreach (var measPin in measPinList)
                        {
                            var forceIPerPin = "";
                            if (isMeasE && measPin.ForceConditions.Count > 1)
                            {
                                var forcePin1 = measPin.ForceConditions[0].ForcePins
                                    .Find(a => a.PinName == measPin.PinName); //First condition
                                forceIe1 += DataConvertor.ConvertForceValueToGlbSpec(forcePin1) + ",";
                                var forcePin2 = measPin.ForceConditions[1].ForcePins
                                    .Find(a => a.PinName == measPin.PinName); //Second condition
                                forceIe2 += DataConvertor.ConvertForceValueToGlbSpec(forcePin2) + ",";
                            }
                            else if (measPin.ForceConditions.Count > 0)
                            {
                                if (measPin.MeasType == "MeasR2") //MeasR2 only need one force condition
                                {
                                }

                                foreach (var condition in measPin.ForceConditions)
                                foreach (var forcePin in condition.ForcePins)
                                {
                                    var measPinName = measPin.PinName.Split(':').Length == 2
                                        ? measPin.PinName.Split(':')[1]
                                        : measPin.PinName;
                                    if (!IsMeasPinInForcePin(forcePin.PinName, measPinName) ||
                                        forcePin.Type == ForceConditionType.Others)
                                        continue;
                                    if (forcePin.ForceType == "I")
                                    {
                                        if (forcePin.ForceJob == "")
                                            forceIPerPin +=
                                                DataConvertor.ConvertForceValueToGlbSpec(forcePin) + forceDelimiter;
                                        else
                                            forceIPerPin +=
                                                forcePin.ForceJob + ":" +
                                                DataConvertor.ConvertForceValueToGlbSpec(forcePin) + ",";
                                    }
                                }
                            }

                            forceI += forceIPerPin.Trim(',') + ",";
                        }

                        forceI = measPinList.Count > 0 && forceI.Length > 0
                            ? forceI.Remove(forceI.Length - 1, 1)
                            : forceI;
                        forceIe1 = forceIe1.Trim(',');
                        forceIe2 = forceIe2.Trim(',');
                        forceIe1 += "+";
                        forceIe2 += "+";
                        forceI += "&";
                    }

                    //if (HardIpDataMain.ConfigData.SeqDelimiter == "|")
                    forceI = forceI.Replace("&", "|");
                    //else if (!forceI.Contains("+")) forceI = forceI.Replace("&", "+");

                    if (isMeasE)
                        forceI = forceIe1.Remove(forceIe1.Length - 1, 1) + "_" +
                                 forceIe2.Remove(forceIe2.Length - 1, 1);
                    else
                        forceI = forceI.Remove(forceI.Length - 1, 1);

                    return forceI;
                }

                return "";
            }

            return "";
        }

        public static string GetIRange(HardIpPattern pattern, HardIpReference info, string voltage)
        {
            var currentRange = "";
            var sequenceList = new List<string>();
            var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount > 0)
            {
                for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    var currentRangeList = new List<string>();

                    var measPinList = pattern.MeasPins.Where(a =>
                        a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC" &&
                        !a.PinName.StartsWith("FT", StringComparison.OrdinalIgnoreCase)).ToList();
                    foreach (var measPin in measPinList)
                    {
                        var range = measPin.GetCurrentRangeByVoltage(voltage);

                        //if (!measPin.PinName.Contains("::") && measPin.PinName.Contains(":"))
                        //    range = measPin.PinName.Split(':')[0] + ":" + range;

                        currentRangeList.Add(range);
                        if (measPin.PinName.Contains("::"))
                            currentRangeList.Add(range);
                    }

                    sequenceList.Add(string.Join(",", currentRangeList));
                }

                currentRange = string.Join("+", sequenceList);
                //if (!currentRange.Contains("+")) currentRange = currentRange.Replace("&", "+");
            }

            if (pattern.MeasPins.Exists(s =>
                Regex.IsMatch(s.MeasType, @"^(measi|measidiff|(MeasR[1|2]))", RegexOptions.IgnoreCase)))
                currentRange = GetIRangeByJob(currentRange);
            if (currentRange == "0")
            {
            }

            return DataConvertor.RemoveDummy(currentRange, @"\+|,");
        }

        public static string GetPrePat(HardIpPattern pattern)
        {
            var interposePrePat = "";
            if (pattern.ForceConditionList.Count > 0)
                foreach (var condition in pattern.ForceConditionList)
                foreach (var pin in condition.ForcePins)
                {
                    if (pin.Type == ForceConditionType.Normal)
                    {
                        if (pin.ForceJob == "")
                            interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                               DataConvertor.ConvertForceValueToGlbSpec(pin) + ";";
                        else
                            interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                               DataConvertor.ConvertForceValueToGlbSpec(pin) + ":" + pin.ForceJob + ";";
                    }

                    if (pin.Type == ForceConditionType.Others)
                        interposePrePat += pin.PinName + ":" + pin.ForceValue + ";";
                }

            #region insert SweepVoltage and SweepCodes information

            if (pattern.SweepVoltage.Count >= 1)
            {
                var interposePrePats = new List<string>();
                foreach (var sweepVoltage in pattern.SweepVoltage)
                {
                    var srcCodeIndx = sweepVoltage.Key.Equals("X", StringComparison.CurrentCultureIgnoreCase)
                        ? "SrcCodeIndx"
                        : "SrcCodeIndxY";
                    foreach (var item in sweepVoltage.Value)
                    {
                        var data = new SweepVData(item);
                        var one = data.PinName + ":V:" + data.Start + "+[" + srcCodeIndx + "]*" + data.Step;
                        interposePrePats.Add(one);
                    }
                }

                interposePrePat = string.Join(";", interposePrePats);
            }

            if (pattern.SweepCodes.Count >= 1) interposePrePat += "Sweep_Name:SrcCodeIndx;";

            #endregion

            #region MiscInfo: USL, LSL

            var miscStr = pattern.MiscInfo;
            foreach (var param in miscStr.Split(';'))
            {
                if (!Regex.IsMatch(param, "^USL:" + "|" + "^LSL:", RegexOptions.IgnoreCase))
                    continue;

                interposePrePat += param.Trim() + ";";
            }

            #endregion

            return interposePrePat.Trim(';');
        }

        public static string GetForceTypeByMeasType(string measPinType)
        {
            if (measPinType == MeasType.MeasV)
                return "^I";
            if (measPinType == MeasType.MeasI)
                return "^V";
            if (measPinType == MeasType.MeasR1 || measPinType == MeasType.MeasR2)
                return "^(V|I)$";
            if (measPinType == MeasType.MeasVdiff2) // force condition was Vp|Vn
                return "^(V1P|V2P|V1N|V2N)$";
            return "^Not_Define$";
        }

        private static List<ForcePin> DivideGroupPin(List<ForcePin> forcePinList)
        {
            var newForcePinList = new List<ForcePin>();
            foreach (var forcePin in forcePinList)
            {
                var pinNames = DecomposeGroups(forcePin.PinName);
                foreach (var pinName in pinNames)
                {
                    var newForcePin = forcePin.DeepClone();
                    newForcePin.PinName = pinName;
                    newForcePinList.Add(newForcePin);
                }
            }

            return newForcePinList;
        }

        private static List<ForcePin> GetForcePinForPreMeasure(List<ForcePin> forcePins, List<MeasPin> measPins,
            string measTypeComparison, string blockForceType, List<ForcePin> allForcePins)
        {
            var isMeasE = measPins.Any(pin => pin.MeasType.Equals("measE", StringComparison.CurrentCultureIgnoreCase));
            var newAllForcePins = new List<ForcePin>();
            //MeasE: Measure IO voltage with 2 different force current and calculate the difference to get calibrated voltage measurement (eg. one by one)
            if (isMeasE)
            {
                // The type of force and measure are the same => the force and measure can not the same (eg. measI & force is I)
                newAllForcePins.AddRange(forcePins.FindAll(forcePin =>
                    forcePin.ForceType.Equals(measTypeComparison, StringComparison.OrdinalIgnoreCase) &&
                    !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName)) && !measPins.Exists(g =>
                        g.PinName.Split(':').ToList().Exists(a => a.Equals(forcePin.PinName)))));
                // The type of force pin is not the opposite type of meas pin (eg. measI & force is not V or I)
                newAllForcePins.AddRange(forcePins.FindAll(forcePin =>
                    !Regex.IsMatch(forcePin.ForceType, blockForceType, RegexOptions.IgnoreCase) &&
                    !forcePin.ForceType.Equals(measTypeComparison, StringComparison.OrdinalIgnoreCase) &&
                    !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName))));
                // The type of force pin is the opposite type of meas pin (eg. measI & force is V)
                newAllForcePins.AddRange(forcePins.FindAll(forcePin =>
                    Regex.IsMatch(forcePin.ForceType, blockForceType, RegexOptions.IgnoreCase) &&
                    !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName)) && !measPins.Exists(g =>
                        g.PinName.Split(':').ToList().Exists(a => a.Equals(forcePin.PinName)))));
            }
            else
            {
                foreach (var forcePin in forcePins)
                    if (!allForcePins.Exists(s => s.Equals(forcePin)))
                    {
                        var isMeasTypeComparison =
                            forcePin.ForceType.Equals(measTypeComparison, StringComparison.OrdinalIgnoreCase);
                        var isBlockForceType =
                            Regex.IsMatch(forcePin.ForceType, blockForceType, RegexOptions.IgnoreCase);
                        var isMeasPinExists = measPins.Exists(g =>
                            g.PinName.Split(':').ToList().Exists(a => a.Equals(forcePin.PinName)));

                        // The type of force and measure are the same => the force and measure can not the same (eg. measI & force is I)
                        if (isMeasTypeComparison && !isMeasPinExists)
                            newAllForcePins.Add(forcePin);
                        // The type of force pin is not the opposite type of meas pin (eg. measI & force is not V or I)
                        else if (!isBlockForceType && !isMeasTypeComparison)
                            newAllForcePins.Add(forcePin);
                        // The type of force pin is the opposite type of meas pin (eg. measI & force is V)
                        else if (isBlockForceType && !isMeasPinExists) newAllForcePins.Add(forcePin);
                    }
            }

            return newAllForcePins;
        }

        public static string GetPreMeas(HardIpPattern pattern)
        {
            var preMeasDelimiter = "|";
            if (preMeasDelimiter == "") preMeasDelimiter = "&";
            var info = GetHardIpInfo(pattern);
            var infoSequence = info.MeasSeqStr == "" ? GetMeasStrByPlan(pattern) : info.MeasSeqStr.Split(',').ToList();
            var preMeas = "";
            if (infoSequence.Count > 0)
            {
                for (var index = 1; index <= infoSequence.Count; index++)
                {
                    var measPins = pattern.MeasPins.Where(a => a.SequenceIndex == index).ToList();
                    var allForcePins = new List<ForcePin>();
                    //Collect all force pins in the same sequenceIndex
                    if (!(pattern.FunctionName.Equals(VbtFunctionLib.VifName, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLib.VirName, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLib.VdiffFunc, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLib.LcdMeas, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLib.DvdcTrim, StringComparison.OrdinalIgnoreCase)))
                    {
                        allForcePins.AddRange(measPins
                            .Where(p => !p.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase))
                            .SelectMany(x => x.ForceConditions).SelectMany(x => x.ForcePins).Where(forcePin =>
                                !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName))));
                    }
                    else
                    {
                        if (pattern.OriMeasPins.Count == 1 && CheckByOriMeasPins(pattern, allForcePins))
                        {
                        }
                        else
                        {
                            foreach (var measPin in measPins)
                            {
                                if (measPin.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase))
                                    continue;
                                var blockForceType = GetForceTypeByMeasType(measPin.MeasType);
                                var measTypeComparision = measPin.MeasType.Replace("Meas", "");
                                foreach (var condition in measPin.ForceConditions)
                                foreach (var forcePin in condition.ForcePins)
                                {
                                    var newForcePins = new List<ForcePin> {forcePin.DeepClone()};
                                    var allForcePinsBeforeDecompose = GetForcePinForPreMeasure(newForcePins, measPins,
                                        measTypeComparision, blockForceType, allForcePins);

                                    // Check if force pin is group pin
                                    if (allForcePinsBeforeDecompose != null)
                                        if (allForcePinsBeforeDecompose.Count > 0)
                                        {
                                            var decomposeForcePins = DivideGroupPin(allForcePinsBeforeDecompose);
                                            var allForcePinsAfterDecompose =
                                                GetForcePinForPreMeasure(decomposeForcePins, measPins,
                                                    measTypeComparision, blockForceType, allForcePins);
                                            allForcePins.AddRange(
                                                allForcePinsAfterDecompose.Count != decomposeForcePins.Count
                                                    ? allForcePinsAfterDecompose
                                                    : allForcePinsBeforeDecompose);
                                        }
                                }
                            }
                        }
                    }

                    if (allForcePins.Count > 0)
                    {
                        var disableFrcPins = allForcePins
                            .Where(p => p.ForceType.Equals("DISABLE_FRC", StringComparison.OrdinalIgnoreCase))
                            .GroupBy(p => p.PinName).ToDictionary(p => p.Key, p => p.ToList());
                        foreach (var disableFrcPin in disableFrcPins)
                        {
                            if (disableFrcPin.Value.Any(p => string.IsNullOrEmpty(p.ForceValue))) continue;
                            disableFrcPin.Value.ForEach(p => p.ForceType = "V");
                            var groupPins = DecomposeGroups(disableFrcPin.Key);
                            groupPins.RemoveAll(p => Regex.IsMatch(p, "REFCLK", RegexOptions.IgnoreCase));
                            allForcePins.Insert(0,
                                new ForcePin
                                {
                                    PinName = disableFrcPin.Key, ForceValue = "DISABLE_FRC",
                                    Type = ForceConditionType.Others
                                });
                            if (groupPins.Count == disableFrcPin.Value.Count)
                            {
                                var interFrcPinIndex = 0;
                                foreach (var forceFrc in disableFrcPin.Value)
                                {
                                    forceFrc.PinName = groupPins[interFrcPinIndex];
                                    interFrcPinIndex++;
                                }
                            }
                        }

                        foreach (var forcePin in allForcePins)
                        {
                            if (forcePin.Type == ForceConditionType.Normal)
                            {
                                if (forcePin.ForceJob == "")
                                    preMeas += forcePin.PinName + ":" + forcePin.ForceType + ":" +
                                               DataConvertor.ConvertForceValueToGlbSpec(forcePin) + ";";
                                else
                                    preMeas += forcePin.PinName + ":" + forcePin.ForceType + ":" +
                                               DataConvertor.ConvertForceValueToGlbSpec(forcePin) + ":" +
                                               forcePin.ForceJob + ";";
                            }

                            if (forcePin.Type == ForceConditionType.Others)
                            {
                                if (forcePin.ForceJob == "")
                                    preMeas += forcePin.PinName + ":" + forcePin.ForceValue + ";";
                                else
                                    preMeas += forcePin.PinName + ":" + forcePin.ForceValue + ":" + forcePin.ForceJob +
                                               ";";
                            }
                        }
                    }

                    preMeas = preMeas.Trim(';') + preMeasDelimiter;
                }

                if (preMeas.Trim(preMeasDelimiter[0]) != "")
                    return preMeas.Substring(0, preMeas.Length - 1);
            }

            return "";
        }

        private static bool CheckByOriMeasPins(HardIpPattern pattern, List<ForcePin> allForcePins)
        {
            var newAllForcePins = new List<ForcePin>();
            foreach (var measPin in pattern.OriMeasPins)
            {
                var blockForceType = GetForceTypeByMeasType(measPin.MeasType);
                var measTypeComparision = measPin.MeasType.Replace("Meas", "");
                foreach (var condition in measPin.ForceConditions)
                foreach (var forcePin in condition.ForcePins)
                {
                    var newForcePins = new List<ForcePin> {forcePin.DeepClone()};
                    var allForcePinsBeforeDecompose = GetForcePinForPreMeasure(newForcePins, pattern.OriMeasPins,
                        measTypeComparision, blockForceType, allForcePins);
                    newAllForcePins.AddRange(allForcePinsBeforeDecompose);
                }
            }

            return !newAllForcePins.Any();
        }

        public static string GetMeasSequence(HardIpPattern pattern)
        {
            var info = GetHardIpInfo(pattern);
            var measList = new Dictionary<int, int>();
            foreach (var pin in pattern.MeasPins)
            {
                var maxForceCnt = 1;
                if (pin.ForceConditions.Any())
                {
                    if (pin.MeasType == MeasType.MeasI || pin.MeasType == MeasType.MeasV)
                        maxForceCnt = pin.ForceConditions.SelectMany(x => x.ForcePins).Max(x => x.ForceCnt);
                    else if (pin.MeasType == MeasType.MeasR2)
                        maxForceCnt = pin.ForceConditions.SelectMany(x => x.ForcePins).Max(x => x.ForceCnt) / 2;
                    maxForceCnt = maxForceCnt < 1 ? 1 : maxForceCnt;
                }

                if (pin.MeasType != "MeasC" && pin.MeasType != MeasType.MeasLimit &&
                    pin.MeasType != MeasType.MeasCalc && !measList.ContainsKey(pin.SequenceIndex))
                    measList.Add(pin.SequenceIndex, maxForceCnt);
            }

            var infoSequence = info.MeasSeqStr == "" ? GetMeasStrByPlan(pattern) : info.MeasSeqStr.Split(',').ToList();
            for (var i = 1; i <= infoSequence.Count; i++)
            {
                var mySeq = infoSequence[i - 1];
                if (!measList.ContainsKey(i))
                {
                    infoSequence[i - 1] = "N";
                }
                else
                {
                    mySeq = new StringBuilder().Insert(0, mySeq, measList[i]).ToString();
                    infoSequence[i - 1] = mySeq;
                }
            }

            return string.Join(",", infoSequence.ToArray()).ToUpper();
        }

        public static List<string> GetSeqLstFromMiscInfo(string miscInfo)
        {
            var testSeqSetting = miscInfo.Split(';').ToList().Find(s =>
                Regex.IsMatch(s, HardIpConstData.RegTestSequence, RegexOptions.IgnoreCase));
            if (testSeqSetting == null)
                return null;
            return testSeqSetting.Split(':')[1].Split(',').ToList();
        }

        public static bool IsValidPatName(string patternName)
        {
            return Regex.IsMatch(patternName, @"^((dd_)|(cz_)|(pp_)|(mn_)|(ht_)).*", RegexOptions.IgnoreCase) ||
                   patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                   Regex.IsMatch(patternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase);
        }

        public static string GetByPair(List<string> list, string delimiter1, string delimiter2, out bool errorFlag)
        {
            var data = "";
            var isOdd = true;
            foreach (var item in list)
                if (isOdd)
                {
                    data += item + delimiter1;
                    isOdd = false;
                }
                else
                {
                    data += item + delimiter2;
                    isOdd = true;
                }

            errorFlag = !isOdd;
            return data.Length > 0 ? data.Remove(data.Length - 1, 1) : data;
        }

        public static string GetIRangeByJob(string measureIRange)
        {
            var measList = Regex.Split(measureIRange, @"[,+]");
            var delimiter = measureIRange.ToCharArray().Where(s => s == ',' || s == '+').ToList();
            var jobList = new List<string>();
            string[] jobValue;

            // Get array size and job list
            var measCnt = measList.Length;
            var jobCnt = 0;
            foreach (var pin in measList)
            {
                jobValue = pin.Split(';');
                jobCnt = jobValue.Length > jobCnt ? jobValue.Length : jobCnt;
                foreach (var job in jobValue)
                    if (job.Split(':').Length > 1)
                        jobList.Add(job.Split(':')[0]);
            }

            if (jobList.Count > 1)
            {
                //Transfer data to array
                var array = new string[measCnt, jobCnt];
                for (var i = 0; i < measCnt; i++)
                {
                    jobValue = measList[i].Split(';');
                    for (var j = 0; j < jobValue.Length; j++)
                        array[i, j] = jobValue[j].Split(':').Length > 1
                            ? jobValue[j].Split(':')[1]
                            : jobValue[j].Split(':')[0];
                }

                //Convert to new IRange string
                var range = "";
                for (var j = 0; j < jobCnt; j++)
                {
                    range += jobList[j] + "=";
                    for (var i = 0; i < measCnt; i++)
                        if (measList[i].Split(';').Length > 1)
                            range += array[i, j] + (i >= delimiter.Count ? "" : delimiter[i].ToString());
                        else
                            range += array[i, 0] + (i >= delimiter.Count ? "" : delimiter[i].ToString());
                    range += ";";
                }

                range = range.Remove(range.Length - 1, 1);

                return range;
            }

            return measureIRange;
        }

        public static bool IsAllMeasI(HardIpPattern pattern)
        {
            return pattern.MeasPins.All(p => p.MeasType == MeasType.MeasI);
        }

        public static bool IsAllTheSameType(List<MeasPin> measPins, string measType)
        {
            return measPins.All(p => p.MeasType == measType);
        }

        public static void ProcessMeasPinTName(HardIpPattern pattern)
        {
            var patInfo = GetHardIpInfo(pattern);
            var nameUseType = DetermineTNameUseType(HardIpDataMain.ConfigData.NameConflictUse);
            var tNameMeasC = GetCusStrDigCapData(pattern);
            if (pattern.ExtraPattern != null)
            {
                var eyeInfo = GetHardIpInfo(pattern.ExtraPattern.Pattern.GetLastPayload());
                patInfo.MeasName = patInfo.MeasName + eyeInfo.MeasName;
                patInfo.SeqInfo.AddRange(eyeInfo.SeqInfo);
                tNameMeasC = tNameMeasC + GetCusStrDigCapData(pattern.ExtraPattern);
            }

            var pinIndex = 1;
            if (HardIpDataMain.TestPlanData.PlanHeaderIdx.ContainsKey(pattern.SheetName))
                pinIndex = HardIpDataMain.TestPlanData.PlanHeaderIdx[pattern.SheetName]["measIndex"];
            var seqCount = patInfo.SeqInfo.Count > 0 ? patInfo.SeqInfo.Count : pattern.TestPlanSequences.Count;
            //search all valid test name from pattern or test plan, plan priority first
            List<MeasPin> measPins;
            //MeasType !=MeasC
            for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var tNamePatInfo = "";

                if (patInfo.MeasName.Split('+').Length >= sequenceIndex &&
                    patInfo.MeasName.Split('+')[sequenceIndex - 1] != "")
                    //T_MeasName = T_MeasName + patInfo.MeasName.Split('+')[sequenceIndex - 1] + "+";
                    tNamePatInfo = patInfo.MeasName.Split('+')[sequenceIndex - 1];
                measPins = pattern.MeasPins.Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != MeasType.MeasC)
                    .ToList();
                foreach (var measPin in measPins)
                {
                    if (nameUseType == NameType.TestPlanOnly) break; //only use test plan item
                    if (measPin.TestName == "" && nameUseType != NameType.PatInfoOnly)
                        measPin.TestName = tNamePatInfo;
                    if (measPin.TestName != tNamePatInfo && tNamePatInfo != "")
                    {
                        var errorMessage = string.Format("Conflict TName between PatInfo:{0} and TestPlan:{1}",
                            tNamePatInfo, measPin.TestName);
                        EpplusErrorManager.AddError(HardIpErrorType.ConflictTName, ErrorLevel.Warning,
                            pattern.SheetName, measPin.RowNum, pinIndex, errorMessage);
                        if (nameUseType == NameType.ConflictUseInfo)
                            measPin.TestName = tNamePatInfo;
                    }
                }
            }

            //MeasType == MeasC
            measPins = pattern.MeasPins.Where(a => a.MeasType == MeasType.MeasC).ToList();
            if (tNameMeasC.Contains("DSSC_OUT"))
            {
                tNameMeasC = tNameMeasC.Replace("DSSC_OUT,", "").Trim(',');
                var capSet = tNameMeasC.Split(',');
                var i = 0;
                foreach (var measPin in measPins)
                {
                    if (nameUseType == NameType.TestPlanOnly) break; //only use test plan item
                    if (i < capSet.Length && measPin.TestName == "") measPin.TestName = capSet[i].Split(':')[1];
                    if (i < capSet.Length && measPin.TestName != capSet[i].Split(':')[1])
                    {
                        var errorMessage = string.Format("Conflict TName between PatInfo:{0} and TestPlan:{1}",
                            capSet[i].Split(':')[1], measPin.TestName);
                        EpplusErrorManager.AddError(HardIpErrorType.ConflictTName, ErrorLevel.Warning,
                            pattern.SheetName, measPin.RowNum, pinIndex, errorMessage);
                        if (nameUseType == NameType.ConflictUseInfo)
                            measPin.TestName = capSet[i].Split(':')[1];
                    }

                    i++;
                }
            }
        }

        private static NameType DetermineTNameUseType(string type)
        {
            if (type.Equals("patInfo", StringComparison.CurrentCulture))
                return NameType.ConflictUseInfo;
            if (type.Equals("testPlan", StringComparison.CurrentCulture))
                return NameType.ConflictUsePlan;
            if (type.Equals("patInfoOnly", StringComparison.CurrentCulture))
                return NameType.PatInfoOnly;
            if (type.Equals("testPlanOnly", StringComparison.CurrentCulture))
                return NameType.TestPlanOnly;
            return NameType.ConflictUseInfo;
        }

        public static string CheckInfoByStoreName(string info, string storeName, char sign, bool isTestSeq = false)
        {
            if (string.IsNullOrEmpty(storeName)) return info;
            try
            {
                var result = new List<string>();
                var storeNameList = storeName.Split('+').ToList();
                if (isTestSeq)
                {
                    //if (info.Split(',').Count() != storeNameList.Count) return info;
                    var i = 0;
                    if (info.Split(',').Length == 1)
                    {
                        var seqTmp = info;
                        for (var k = 0; k < storeNameList[i].Split(':').Length - 1; k++) seqTmp = seqTmp + info;
                        result.Add(seqTmp);
                    }
                    else
                    {
                        foreach (var seq in info.Split(','))
                        {
                            var seqTmp = seq;
                            for (var k = 0; k < storeNameList[i].Split(':').Length - 1; k++) seqTmp = seqTmp + seq;
                            result.Add(seqTmp);
                            i++;
                        }
                    }

                    return string.Join(",", result);
                }
                else
                {
                    //if (info.Split(sign).Count() != storeNameList.Count) return info;

                    var i = 0;
                    if (info.Split(sign).Length == 1)
                        result.Add(storeNameList[i].Contains(":") ? info.Replace(",", ":") : info);
                    else
                        foreach (var seqInfo in info.Split(sign))
                        {
                            result.Add(storeNameList[i].Contains(":") ? seqInfo.Replace(",", ":") : seqInfo);
                            i++;
                        }

                    return string.Join(sign.ToString(), result);
                }
            }
            catch (Exception e)
            {
                Response.Report(e.ToString(), MessageLevel.Error, 0);
            }

            return info;
        }

        public static string GetInstrumentInfo(HardIpPattern pattern, string item)
        {
            var itemSeq = new List<string>();
            var regItem = string.Format(@"{0}\s*=\s*(?<value>[\w,#]+)", item);
            for (var seqIndex = 1; seqIndex <= pattern.MeasPins.Max(p => p.SequenceIndex); seqIndex++)
            {
                var index = seqIndex;
                var pins = pattern.MeasPins.Where(p => index == p.SequenceIndex);
                var setups = pins.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase))
                    .GroupBy(p => p.RfInstrumentSetup).ToDictionary(p => p.Key, p => p.ToList());
                var items = new List<string>();
                foreach (var setup in setups.Keys)
                {
                    var selectItem = "";
                    foreach (var setupInfo in setup.Split('$'))
                        if (Regex.IsMatch(setupInfo, regItem, RegexOptions.IgnoreCase))
                        {
                            selectItem = Regex.Match(setupInfo, regItem, RegexOptions.IgnoreCase).Groups["value"]
                                .ToString();
                            break;
                        }

                    items.Add(selectItem);
                }

                itemSeq.Add(string.Join("+", items));
            }

            return string.Join("|", itemSeq);
        }

        private enum NameType
        {
            PatInfoOnly,
            TestPlanOnly,
            ConflictUseInfo,
            ConflictUsePlan
        }
    }
}