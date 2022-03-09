using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.NonIgxlSheets;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.GenerateIgxl.HardIp.DividerManager.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.DividerManager
{
    public class DividerMain
    {
        private static HardIpReference _patInfo;
        private static bool _isMeasE;
        private static bool _isMeasVdiff2;
        private static bool _isMeasI2;
        private static bool _isMeasVdiff;
        private static int _measCount;

        public static List<HardIpPattern> DivideMeasPins(HardIpPattern pattern, bool isHardIpUniversal,
            bool isFlowUse = false)
        {
            var resultList = new List<HardIpPattern>();
            _patInfo = SearchInfo.GetHardIpInfo(pattern);
            SearchInfo.InitialMeasC(pattern, _patInfo);
            _isMeasE = SearchInfo.IsMeasE(pattern);
            _measCount = SearchInfo.MeasECount(pattern);
            _isMeasVdiff2 = SearchInfo.IsMeasVdiff2(pattern);
            _isMeasVdiff = SearchInfo.IsMeasVdiff(pattern);
            _isMeasI2 = SearchInfo.IsMeasI2(pattern);
            var isRepeatLimit = SearchInfo.IsRepeatLimit(pattern.MiscInfo);

            var totalPins = new List<MeasPin>();
            var patternVir = new HardIpPattern();
            var patternFreq = new HardIpPattern();
            //bool isMeasVdiff = SearchInfo.IsMeasVdiff(pattern);
            var otherPins = new List<MeasPin>();

            foreach (var pin in pattern.MeasPins)
                if (pin.MeasType == MeasType.MeasCalc || pin.MeasType == MeasType.MeasLimit ||
                    pin.MeasType == MeasType.MeasC || pin.MeasType == MeasType.MeasN || pin.MeasType == "")
                    otherPins.Add(pin);

            #region Test plan does not have any Meas Pins, refer to pattern info. If pattern info contains src or capture information, choose VFI as VBT. otherwise, Function T.

            var isUseTestPlan = false;
            if (_patInfo.SeqInfo.Count == 0)
            {
                isUseTestPlan = true;
                if (pattern.MeasPins.Count == 0)
                {
                    if (isHardIpUniversal)
                    {
                        if (pattern.IsNonHardIpBlock || string.IsNullOrEmpty(_patInfo.SendBitStr) &&
                            string.IsNullOrEmpty(_patInfo.CapBitStr))
                            pattern.FunctionName = VbtFunctionLib.FunctionalTUpdated; //"Functional_T_Updated";
                        else
                            pattern.FunctionName = VbtFunctionLib.VifName; // "Meas_FreqVoltCurr_Universal_func";
                    }

                    resultList.Add(pattern);
                    return resultList;
                }

                RemoveMeasPinsByMeasType(pattern);
                totalPins.AddRange(pattern.MeasPins);
            }
            else
            {
                if (_patInfo.NewInfo != null)
                    foreach (var info in _patInfo.NewInfo.SeqInfo)
                        totalPins.AddRange(info.MeasPins);
                else
                    foreach (var info in _patInfo.SeqInfo)
                    foreach (var measPin in info.MeasPins)
                    {
                        if (measPin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                        {
                            var pnPins = Regex.Split(measPin.PinName, "::").ToList();
                            pnPins.Sort();
                            foreach (var pin in pnPins)
                            {
                                var copySeqPin = new MeasPin();
                                copySeqPin.Copy(measPin);
                                copySeqPin.PinName = pin;
                                copySeqPin.MeasType = MeasType.MeasV;
                                copySeqPin.SequenceIndex = measPin.SequenceIndex;
                                totalPins.Add(copySeqPin);
                            }
                        }

                        totalPins.Add(measPin);
                    }
            }

            #endregion

            #region Align measpins with VFI and VIR

            var fPins = totalPins.Where(p =>
                p.MeasType.Equals(MeasType.MeasF) || p.MeasType.Equals(MeasType.MeasFdiff) ||
                p.MeasType.Equals(MeasType.MeasN)).ToList();
            var rPins = ProjectConfigSingleton.Instance().GetProjectConfigValue("HardIP", "UniversalVBT") == "VFIOnly"
                ? new List<MeasPin>()
                : totalPins.Where(p =>
                    p.MeasType.Equals(MeasType.MeasIdiff) || p.MeasType.Equals(MeasType.MeasVdiff) ||
                    p.MeasType.Equals(MeasType.MeasE) || p.MeasType.Equals(MeasType.MeasR1) ||
                    p.MeasType.Equals(MeasType.MeasR2)).ToList();
            var virFlag = rPins.Count > 0 || _isMeasE;
            var freqFlag = fPins.Count > 0;
            var wiPins = totalPins.Where(p => p.MeasType.Equals(MeasType.WiSrc) || p.MeasType.Equals(MeasType.WiMeas))
                .ToList();
            if (isHardIpUniversal)
            {
                if (virFlag)
                {
                    if (!freqFlag)
                    {
                        rPins = totalPins;
                        patternVir.MeasPins.AddRange(otherPins);
                    }
                    else
                    {
                        totalPins.RemoveAll(p => rPins.Contains(p));
                        fPins = totalPins;
                    }
                }
                else
                {
                    fPins = totalPins;
                }
            }

            #endregion

            totalPins.RemoveAll(otherPins.Contains);
            if (!isHardIpUniversal)
            {
                var customVbtPattern = new HardIpPattern();
                if (!pattern.FunctionName.Equals(VbtFunctionLib.RfFunc) &&
                    !pattern.FunctionName.Equals(VbtFunctionLib.LcdMeas))
                {
                    var sortedPins = DataConvertor.SortMeasPin(totalPins);
                    sortedPins =
                        ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _measCount, isUseTestPlan, isFlowUse);
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);
                    customVbtPattern.Copy(pattern);
                    customVbtPattern.MeasPins = sortedPins;
                    customVbtPattern.FunctionName = pattern.FunctionName;
                    customVbtPattern.MeasPins.AddRange(otherPins);
                    DividerCommonLogic.RemoveIgnoredSequence(customVbtPattern);
                    resultList.Add(customVbtPattern);
                }
                else
                {
                    resultList.Add(pattern);
                }

                return resultList;
            }

            if (!_isMeasVdiff2)
            {
                #region BB RF part

                if (wiPins.Count > 0)
                {
                    var hardIpPattern = new HardIpPattern();
                    var sortedPins = DataConvertor.SortMeasPin(wiPins);
                    sortedPins =
                        ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _measCount, isUseTestPlan, isFlowUse);
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);
                    hardIpPattern.Copy(pattern);
                    hardIpPattern.MeasPins = sortedPins;
                    hardIpPattern.FunctionName = pattern.FunctionName;

                    DividerCommonLogic.RemoveIgnoredSequence(hardIpPattern);
                    resultList.Add(hardIpPattern);
                    return resultList;
                }

                if (rPins.Count > 0)
                {
                    var sortedPins = DataConvertor.SortMeasPin(rPins);
                    sortedPins =
                        ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _measCount, isUseTestPlan, isFlowUse);
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);
                    patternVir.Copy(pattern);

                    patternVir.MeasPins = sortedPins;
                    patternVir.FunctionName = VbtFunctionLib.VirName; //"Meas_VIR_IO_Universal_func";
                    if (fPins.Count == 0)
                        patternVir.MeasPins.AddRange(otherPins);
                    if (freqFlag)
                        RemoveUnusedCalEqn(patternVir);
                    if (_isMeasVdiff)
                        patternVir.SpecialMeasType =
                            pattern.MeasPins.Exists(s =>
                                s.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                                ? MeasType.MeasVocm
                                : MeasType.MeasVdiff;
                    if (_isMeasI2)
                        patternVir.SpecialMeasType = MeasType.MeasI2;
                    DividerCommonLogic.RemoveIgnoredSequence(patternVir);
                    resultList.Add(patternVir);
                }

                #endregion

                #region freq(default use)

                if (fPins.Count > 0)
                {
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, fPins, isRepeatLimit);
                    fPins = ProcessSpecialMeasItems(fPins, _isMeasVdiff2, _measCount, isUseTestPlan, isFlowUse);
                    var sortedPins = DataConvertor.SortMeasPin(fPins);

                    patternFreq.Copy(pattern);

                    patternFreq.MeasPins = sortedPins;
                    patternFreq.FunctionName = VbtFunctionLib.VifName; //"Meas_FreqVoltCurr_Universal_func";
                    RemoveMeasPinsByMeasType(patternVir);
                    if (_isMeasI2)
                        patternFreq.SpecialMeasType = MeasType.MeasI2;
                    if (_isMeasVdiff)
                        patternFreq.SpecialMeasType =
                            pattern.MeasPins.Exists(s =>
                                s.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                                ? MeasType.MeasVocm
                                : MeasType.MeasVdiff;
                    DividerCommonLogic.RemoveIgnoredSequence(patternFreq);
                    patternFreq.MeasPins.AddRange(otherPins);
                    resultList.Add(patternFreq);

                    #endregion

                    if (virFlag && freqFlag)
                    {
                        patternVir.DivideFlag = "_VIR";
                        patternFreq.DivideFlag = "";
                    }
                }
                else if (otherPins.Count > 0 && rPins.Count == 0)
                {
                    patternFreq.Copy(pattern);
                    patternFreq.MeasPins = otherPins;
                    patternFreq.FunctionName = VbtFunctionLib.VifName; //"Meas_FreqVoltCurr_Universal_func";
                    resultList.Add(patternFreq);
                }
            }

            #region Vdiff2 as Meas_Lpdp_Vdiff2_fuc

            else
            {
                var sortedPins = DataConvertor.SortMeasPin(fPins);
                sortedPins = ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _measCount, isUseTestPlan, isFlowUse);
                SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);


                pattern.MeasPins = sortedPins;
                pattern.FunctionName = VbtFunctionLib.VdiffFunc;
                DividerCommonLogic.RemoveIgnoredSequence(pattern);
                pattern.MeasPins.AddRange(otherPins);
                resultList.Add(pattern);
            }

            #endregion

            return resultList;
        }

        private static void RemoveUnusedCalEqn(HardIpPattern pattern)
        {
            const string regPin = @"(?<pinName>[\w]+)[\(][\w]+[\)]";
            var calcEqns = pattern.CalcEqn.Split(';').ToList();
            var cnt = calcEqns.Count;
            if (calcEqns.Count == 0)
                return;
            for (var index = 0; index < cnt; index++)
            {
                var matches = Regex.Matches(calcEqns[index], regPin);
                foreach (Match match in matches)
                {
                    var pinName = match.Groups["pinName"].ToString();
                    if (pattern.FunctionName == VbtFunctionLib.VifName && SearchInfo.GetPinType(pinName) == "I/O")
                    {
                        calcEqns.RemoveAt(index);
                        break;
                    }

                    if (pattern.FunctionName == VbtFunctionLib.VirName && SearchInfo.GetPinType(pinName) != "I/O")
                    {
                        calcEqns.RemoveAt(index);
                        break;
                    }
                }
            }

            pattern.CalcEqn = string.Join(";", calcEqns);
        }

        private static void RemoveMeasPinsByMeasType(HardIpPattern pattern)
        {
            var pinGroups = from pin in pattern.MeasPins group pin by pin.SequenceIndex into g select g;
            foreach (var pinGroup in pinGroups)
            {
                var seqIndex = pinGroup.Key;
                var isIdiff = pinGroup.ToList().Exists(s => s.MeasType == "MeasIdiff");
                var isVdiff = pinGroup.ToList().Exists(s => s.MeasType == "MeasVdiff");
                if (isIdiff) pattern.MeasPins.RemoveAll(s => s.SequenceIndex == seqIndex && s.MeasType != "MeasIdiff");
                if (isVdiff)
                    pattern.MeasPins.RemoveAll(s =>
                        s.SequenceIndex == seqIndex &&
                        !s.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase));
            }

            var measCPins = pattern.MeasPins
                .FindAll(s => s.MeasType == MeasType.MeasC || s.MeasType == MeasType.MeasLimit).ToList();
            pattern.MeasPins.RemoveAll(measCPins.Contains);
            pattern.MeasPins = pattern.MeasPins.OrderBy(x => x.SequenceIndex).ThenBy(x => x.PinName.ToLower()).ToList();
            pattern.MeasPins.AddRange(measCPins);
        }

        private static List<MeasPin> ProcessSpecialMeasItems(List<MeasPin> allPins, bool isVdiff2, int measECount,
            bool isUseTestPlan, bool isFlowUse)
        {
            var resultPins = new List<MeasPin>();
            if (allPins.Count == 0)
                return resultPins;
            var seqCount = allPins.Max(p => p.SequenceIndex);
            for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var specialPins = new List<MeasPin>();
                var seqPins = allPins.Where(p => p.SequenceIndex == sequenceIndex).ToList();
                var isVDiff = seqPins.Exists(p =>
                    p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase));
                var inSeqPins = new List<MeasPin>();
                if (_measCount > 0)
                {
                    if (sequenceIndex > seqCount - measECount)
                        foreach (var seqPin in seqPins)
                            if (seqPins.FirstOrDefault(p => p.MeasType.Equals(MeasType.MeasE)) == null)
                            {
                                var copyPin = new MeasPin();
                                copyPin.Copy(seqPin);
                                copyPin.PinName = seqPin.PinName;
                                copyPin.SequenceIndex = seqPin.SequenceIndex;
                                copyPin.MeasType = MeasType.MeasE;
                                specialPins.Add(seqPin);
                                inSeqPins.Add(copyPin);
                            }
                            else
                            {
                                inSeqPins.Add(seqPin);
                            }

                    if (isFlowUse && !isUseTestPlan)
                    {
                        inSeqPins.InsertRange(0, specialPins);
                        inSeqPins.InsertRange(0, specialPins);
                    }
                }
                else if (isVDiff && !isUseTestPlan && !isFlowUse)
                {
                    //1. P,N 2. Vdiff 3. Vocm

                    foreach (var seqPin in seqPins)
                        if (seqPin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                            inSeqPins.Add(seqPin);
                }
                else if (isVdiff2 && !isUseTestPlan)
                {
                }
                else if (seqPins.Count > 0)
                {
                    inSeqPins.AddRange(seqPins);
                }

                resultPins.AddRange(inSeqPins);
            }

            return resultPins;
        }
    }
}