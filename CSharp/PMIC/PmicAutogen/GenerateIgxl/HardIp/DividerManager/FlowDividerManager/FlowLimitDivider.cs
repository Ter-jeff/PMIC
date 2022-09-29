using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;

namespace PmicAutogen.GenerateIgxl.HardIp.DividerManager.FlowDividerManager
{
    public class FlowLimitDivider
    {
        public static List<HardIpPattern> DivideUseLimit(List<HardIpPattern> patternList)
        {
            if (HardIpDataMain.TestPlanData.ExtendLimits == false)
            {
                foreach (var pattern in patternList)
                    try
                    {
                        var info = SearchInfo.GetHardIpInfo(pattern);
                        SearchInfo.ProcessMeasPinTName(pattern);
                        var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                        var measPins = pattern.MeasPins
                            .Where(a => !a.PinName.StartsWith("FT", StringComparison.OrdinalIgnoreCase)).ToList();
                        var otherLimit = measPins.Where(x => x.SequenceIndex == 0).ToList();
                        if (seqCount > 0)
                            for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                            {
                                var loopCnt = 0;
                                var measPin = measPins.Where(x => x.SequenceIndex == sequenceIndex).ToList();
                                if (measPin.Count > 0 &&
                                    (measPin.All(p => p.MeasType == MeasType.MeasI) ||
                                     measPin.All(p => p.MeasType == MeasType.MeasR1) ||
                                     measPin.All(p => p.MeasType == MeasType.MeasR2)))
                                {
                                    foreach (var pin in measPin)
                                    {
                                        var pinList = pin.ForceConditions.SelectMany(x => x.ForcePins)
                                            .Where(y => y.PinName == pin.PinName).ToList();
                                        loopCnt += pinList.Any() ? pinList.Max(y => y.ForceCnt) : 0;
                                    }

                                    //loopCnt = loopCnt / measPin.Count;
                                    loopCnt = measPin[0].MeasType == MeasType.MeasR2
                                        ? loopCnt / measPin.Count / 2
                                        : loopCnt / measPin.Count;
                                    loopCnt = loopCnt == 0 ? 1 : loopCnt;
                                    for (var i = 0; i < loopCnt; i++)
                                        foreach (var pin in measPin)
                                        {
                                            var measList = measPins.Where(x =>
                                                x.RowNumForMergeMeas == pin.RowNum && x.PinName == pin.PinName &&
                                                x.MeasType == MeasType.MeasLimit).ToList();
                                            if (i >= 1 && measList.Any())
                                            {
                                                var mergeIdx = i - 1;
                                                var newPin = measList.Count > mergeIdx
                                                    ? measList.ElementAt(mergeIdx)
                                                    : measList.ElementAt(measList.Count - 1);
                                                UpdateLimits(pattern.UseLimitsH, newPin, "H");
                                                UpdateLimits(pattern.UseLimitsL, newPin, "L");
                                                UpdateLimits(pattern.UseLimitsN, newPin, "N");
                                            }
                                            else
                                            {
                                                if (pin.TestName.Contains(','))
                                                {
                                                    var newPin = pin.DeepClone();
                                                    if (pin.TestName.Split(',').Length >= i + 1)
                                                    {
                                                        newPin.TestName = pin.TestName.Split(',')[i];
                                                        UpdateLimits(pattern.UseLimitsH, newPin, "H");
                                                        UpdateLimits(pattern.UseLimitsL, newPin, "L");
                                                        UpdateLimits(pattern.UseLimitsN, newPin, "N");
                                                    }
                                                }
                                                else
                                                {
                                                    UpdateLimits(pattern.UseLimitsH, pin, "H");
                                                    UpdateLimits(pattern.UseLimitsL, pin, "L");
                                                    UpdateLimits(pattern.UseLimitsN, pin, "N");
                                                }
                                            }
                                        }
                                }
                                else
                                {
                                    foreach (var pin in measPin)
                                    {
                                        UpdateLimits(pattern.UseLimitsH, pin, "H");
                                        UpdateLimits(pattern.UseLimitsL, pin, "L");
                                        UpdateLimits(pattern.UseLimitsN, pin, "N");
                                    }
                                }
                            }

                        foreach (var pin in otherLimit)
                        {
                            UpdateLimits(pattern.UseLimitsH, pin, "H");
                            UpdateLimits(pattern.UseLimitsL, pin, "L");
                            UpdateLimits(pattern.UseLimitsN, pin, "N");
                        }


                        if (pattern.SheetName.ToUpper().Contains("DCTEST_IDS"))
                            DataConvertor.SortIdsPins(pattern
                                .UseLimitsN); //Sort MeasPins according to CorePower/OtherPower
                        if (pattern.SheetName.ToUpper().Contains("DCTEST_IDS"))
                            DataConvertor.SortIdsPins(pattern
                                .UseLimitsH); //Sort MeasPins according to CorePower/OtherPower
                        if (pattern.SheetName.ToUpper().Contains("DCTEST_IDS"))
                            DataConvertor.SortIdsPins(pattern
                                .UseLimitsL); //Sort MeasPins according to CorePower/OtherPower
                    }
                    catch (Exception e)
                    {
                        Response.Report(e.ToString(), MessageLevel.Error, 0);
                        throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum +
                                            e);
                    }

                return patternList;
            }

            var hardIpPatterns = new List<HardIpPattern>();
            foreach (var pattern in patternList)
            {
                SearchInfo.ProcessMeasPinTName(pattern);
                hardIpPatterns.Add(pattern);
                if (pattern.MeasPins.Count == 0) //If no MeasPin, do not need divide limit 
                    continue;

                pattern.UseLimitsH = GenerateLimitByJob(pattern, "H");
                pattern.UseLimitsL = GenerateLimitByJob(pattern, "L");
                pattern.UseLimitsN = GenerateLimitByJob(pattern, "N");
            }

            return hardIpPatterns;
        }

        private static List<MeasPin> GenerateLimitByJob(HardIpPattern pattern, string voltage)
        {
            var info = SearchInfo.GetHardIpInfo(pattern);
            var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            var measPins = pattern.MeasPins.Where(a => !a.PinName.StartsWith("FT", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var otherLimit = measPins.Where(x => x.SequenceIndex == 0).ToList();
            var newList = new List<MeasPin>();
            var groupLimits = GroupLimits(pattern, voltage);
            try
            {
                foreach (var limit in groupLimits.Keys)
                {
                    var jobList = groupLimits[limit];
                    var values = limit.Split(',').ToList();
                    var index = 0;


                    for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var loopCnt = 0;
                        var measPin = measPins.Where(x => x.SequenceIndex == sequenceIndex).ToList();
                        if (measPin.Count > 0 && (measPin.All(p => p.MeasType == MeasType.MeasI) ||
                                                  measPin.All(p => p.MeasType == MeasType.MeasR1) ||
                                                  measPin.All(p => p.MeasType == MeasType.MeasR2)))
                        {
                            foreach (var pin in measPin)
                            {
                                var pinList = pin.ForceConditions.SelectMany(x => x.ForcePins)
                                    .Where(y => y.PinName == pin.PinName).ToList();
                                loopCnt += pinList.Any() ? pinList.Max(y => y.ForceCnt) : 0;
                            }

                            //loopCnt = loopCnt / measPin.Count;
                            loopCnt = measPin[0].MeasType == MeasType.MeasR2
                                ? loopCnt / measPin.Count / 2
                                : loopCnt / measPin.Count;
                            loopCnt = loopCnt == 0 ? 1 : loopCnt;
                        }
                        else
                        {
                            loopCnt = 1;
                        }

                        for (var i = 0; i < loopCnt; i++)
                        {
                            var tempList = new List<MeasPin>();
                            foreach (var pin in measPin)
                            {
                                //if (pin.MeasType == "MeasC" && pin.MeasLimitsH.Find(a => DataConvertor.ConvertLimit(a.LoLimit) != "" || DataConvertor.ConvertLimit(a.HiLimit) != "") == null)// If its MeasC pin and without any limits, ignore it
                                //    continue;
                                var newPin = new MeasPin();
                                newPin.Copy(pin);
                                newPin.PinName = pin.PinName;
                                newPin.LowLimit = values[index].Split('$')[0];
                                newPin.HighLimit = values[index].Split('$')[1];
                                newPin.Job = string.Join(",", jobList);
                                newPin.SequenceIndex = pin.SequenceIndex;
                                if (CheckMergePower(newPin)
                                ) //"true" means the pin is power merge pin, and mismatch with its job-enable, skip this pin
                                {
                                    index++;
                                    continue;
                                }

                                if (newPin.TestName.Contains(','))
                                    if (pin.TestName.Split(',').Length >= i + 1)
                                        newPin.TestName = pin.TestName.Split(',')[i];

                                tempList.Add(newPin);
                                index++;
                            }

                            if (loopCnt > 1)
                                index = index - measPin.Count;

                            #region Sort Ids pins

                            if (pattern.SheetName.ToUpper().Contains("DCTEST_IDS"))
                                DataConvertor.SortIdsPins(tempList); //Sort MeasPins according to CorePower/OtherPower

                            #endregion

                            newList.AddRange(tempList);
                        }
                    }

                    foreach (var pin in otherLimit)
                    {
                        var newPin = new MeasPin();
                        newPin.Copy(pin);
                        newPin.PinName = pin.PinName;
                        newPin.LowLimit = values[index].Split('$')[0];
                        newPin.HighLimit = values[index].Split('$')[1];
                        newPin.Job = string.Join(",", jobList);
                        newList.Add(newPin);
                        index++;
                    }
                }
            }
            catch (Exception e)
            {
                Response.Report(e.ToString(), MessageLevel.Error, 0);
            }

            return newList;
        }

        private static Dictionary<string, List<string>> GroupLimits(HardIpPattern pattern, string voltage = "")
        {
            var groupLimits = new Dictionary<string, List<string>>();
            //bool allMeasC = true;
            var valueList = new Dictionary<string, string>();
            foreach (var pin in pattern.MeasPins)
            {
                List<MeasLimit> limits;
                switch (voltage)
                {
                    case "":
                        limits = pin.MeasLimitsN;
                        break;
                    case "H":
                        limits = pin.MeasLimitsH;
                        break;
                    case "L":
                        limits = pin.MeasLimitsL;
                        break;
                    case "N":
                        limits = pin.MeasLimitsN;
                        break;
                    default:
                        limits = pin.MeasLimitsN;
                        break;
                }
                //if (pin.MeasType == "MeasC" && limits.Find(a => DataConvertor.ConvertLimit(a.LoLimit) != "" || DataConvertor.ConvertLimit(a.HiLimit) != "") == null)//  DataConvertor.ConvertLimit(pin.Cp1Lo) == "" && DataConvertor.ConvertLimit(pin.Cp1Hi) == "" && DataConvertor.ConvertLimit(pin.Cp2Lo) == "" && DataConvertor.ConvertLimit(pin.Cp2Hi) == "" && DataConvertor.ConvertLimit(pin.Ft3Lo) == "" && DataConvertor.ConvertLimit(pin.Ft3Hi) == "" && DataConvertor.ConvertLimit(pin.Ft1Lo) == "" && DataConvertor.ConvertLimit(pin.Ft1Hi) == "" && DataConvertor.ConvertLimit(pin.Ft2Lo) == "" && DataConvertor.ConvertLimit(pin.Ft2Hi) == "" && DataConvertor.ConvertLimit(pin.HtolLo) == "" && DataConvertor.ConvertLimit(pin.HtolHi) == "" && DataConvertor.ConvertLimit(pin.QaLo) == "" && DataConvertor.ConvertLimit(pin.QaHi) == "")
                //    continue;
                //allMeasC = false;

                #region Initial MeasLimits for the pin which is not specified in TestPlan

                if (limits.Count == 0)
                    foreach (var job in HardIpDataMain.TestPlanData.AllJobs)
                    {
                        var newLimit = new MeasLimit(job);
                        limits.Add(newLimit);
                    }

                #endregion

                foreach (var limit in limits)
                    if (valueList.ContainsKey(limit.JobName))
                        //valueList[limit.JobName] += DataConvertor.ConvertLimit(limit.LoLimit) + "$" + DataConvertor.ConvertLimit(limit.HiLimit) + ",";
                        valueList[limit.JobName] += limit.LoLimit + "$" + limit.HiLimit + ",";
                    else
                        //valueList.Add(limit.JobName, DataConvertor.ConvertLimit(limit.LoLimit) + "$" + DataConvertor.ConvertLimit(limit.HiLimit) + ",");
                        valueList.Add(limit.JobName, limit.LoLimit + "$" + limit.HiLimit + ",");
            }

            //if (allMeasC)
            //    return groupLimits;
            foreach (var group in valueList) GenGroups(group.Value, group.Key, groupLimits);
            return groupLimits;
        }

        private static void GenGroups(string limitStr, string jobName, Dictionary<string, List<string>> groupLimits)
        {
            if (limitStr != "")
            {
                limitStr = limitStr.Remove(limitStr.Length - 1, 1);
                if (groupLimits.ContainsKey(limitStr))
                {
                    groupLimits[limitStr].Add(jobName);
                }
                else
                {
                    var jobList = new List<string>();
                    jobList.Add(jobName);
                    groupLimits.Add(limitStr, jobList);
                }
            }
        }

        private static bool CheckMergePower(MeasPin pin)
        {
            if (!pin.PinName.Contains("VDD") || HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(pin.PinName)
            ) //PowerMerge pins are Power pin(VDD_)
                return false;
            var jobList = pin.Job;
            //if (jobList.Contains("CP") && !jobList.Contains("FT") &&
            //    (!HardIpDataMain.TempResultData.PowerMerge.CpPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("FT=")))
            //    return true;
            //if (jobList.Contains("FT") && !jobList.Contains("CP") && (!HardIpDataMain.TempResultData.PowerMerge.FtPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("CP=")))
            //    return true;
            if (jobList.Contains("FT") && jobList.Contains("CP"))
            {
                //var newJobList = new List<string>();
                //Remove ft jobs for only-CP pin
                //if (!HardIpDataMain.TempResultData.PowerMerge.FtPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("CP="))
                //{
                //    foreach (var job in jobList.Split(','))
                //    {
                //        if (job.Contains("CP"))
                //            newJobList.Add(job);
                //    }
                //}
                ////Remove cp jobs for only-FT pin
                //if (!HardIpDataMain.TempResultData.PowerMerge.CpPowers.ContainsValue(RemoveJobName(pin.PinName)) || pin.PinName.Contains("FT="))
                //{
                //    foreach (var job in jobList.Split(','))
                //    {
                //        if (job.Contains("FT"))
                //            newJobList.Add(job);
                //    }
                //}
                //if (newJobList.Count > 0)
                //    pin.Job = string.Join(",", newJobList);
            }

            return false;
        }

        public static void UpdateLimits(List<MeasPin> limits, MeasPin pin, string type)
        {
            var measPin = new MeasPin();
            if (pin.PinName.Split('=').Length == 2)
            {
                var relatedJobs = HardIpDataMain.TestPlanData.AllJobs.ToList().Where(p =>
                    Regex.IsMatch(p, pin.PinName.Split('=')[0], RegexOptions.IgnoreCase));
                measPin.Job = string.Join(",", relatedJobs);
            }
            else
            {
                measPin.Job = pin.Job;
            }

            measPin.MeasType = pin.MeasType;
            measPin.RowNumForMergeMeas = pin.RowNumForMergeMeas;
            measPin.SequenceIndex = pin.SequenceIndex;
            measPin.ForceConditions = pin.ForceConditions;
            measPin.RowNum = pin.RowNum;
            measPin.TestName = pin.TestName;
            measPin.PinName = pin.PinName;
            measPin.MiscInfo = pin.MiscInfo;
            var measLimits = new List<MeasLimit>();
            switch (type.ToUpper())
            {
                case "H":
                    measLimits = pin.MeasLimitsH;
                    break;
                case "L":
                    measLimits = pin.MeasLimitsL;
                    break;
                case "N":
                    measLimits = pin.MeasLimitsN;
                    break;
            }

            if (measLimits.Count > 0)
            {
                measPin.LowLimit = measLimits[0].LoLimit;
                measPin.HighLimit = measLimits[0].HiLimit;
            }

            if (!(pin.MeasType.Equals(MeasType.WiMeas) || pin.MeasType.Equals(MeasType.MeasWait)))
                limits.Add(measPin);
        }
    }
}