using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility
{
    public class DataConvertor
    {
        public static string Var => "_VAR";

        public static string ConvertUnits(string limitStr)
        {
            if (limitStr.Contains("10^"))
                limitStr = limitStr.Replace("*10^", "E");
            if (limitStr == "" || limitStr.Contains("E") || Regex.IsMatch(limitStr, @"^(\d|\.|-)+$")
            ) //Limit value may be 1.2E-5
                return limitStr;
            if (Regex.IsMatch(limitStr, @"^(\d|\.|-)+(\w)*$"))
            {
                var limitNum = Regex.Match(limitStr, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (limitNum == "0")
                    return limitNum;
                var limitUnit = limitStr.Replace(limitNum, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(limitUnit, "^m.*"))
                    rate = 1 / (double) 1000;
                else if (Regex.IsMatch(limitUnit, "^u.*"))
                    rate = 1 / (double) 1000000;
                else if (Regex.IsMatch(limitUnit, "^n.*"))
                    rate = 1 / (double) 1000000000;
                else if (Regex.IsMatch(limitUnit.ToLower(), "^k.*"))
                    rate = 1000;
                else if (Regex.IsMatch(limitUnit, "^M.*"))
                    rate = 1000000;
                else if (Regex.IsMatch(limitUnit, "^G.*"))
                    rate = 1000000000;
                double value;
                if (double.TryParse(limitNum, out value)) return (value * rate).ToString("G");
            }

            return limitStr;
        }

        public static string ConvertUnits(string limitStr, out string limitUnit, out string limitScale)
        {
            limitUnit = "";
            limitScale = "";
            if (limitStr == "" || limitStr.Contains("E") || Regex.IsMatch(limitStr, @"^(\d|\.|-)+$")
            ) //Limit value may be 1.2E-5
                return limitStr;
            if (Regex.IsMatch(limitStr, @"^(\d|\.|-)+(\w)*$"))
            {
                var limitNum = Regex.Match(limitStr, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (limitNum == "0")
                    return limitNum;
                limitUnit = limitStr.Replace(limitNum, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(limitUnit, "^m.*"))
                {
                    rate = 1 / (double) 1000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleMilli;
                }
                else if (Regex.IsMatch(limitUnit, "^u.*"))
                {
                    rate = 1 / (double) 1000000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleMicro;
                }
                else if (Regex.IsMatch(limitUnit, "^n.*"))
                {
                    rate = 1 / (double) 1000000000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleNano;
                }
                else if (Regex.IsMatch(limitUnit.ToLower(), "^k.*"))
                {
                    rate = 1000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleKilo;
                }
                else if (Regex.IsMatch(limitUnit, "^M.*"))
                {
                    rate = 1000000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleMega;
                }
                else if (Regex.IsMatch(limitUnit, "^G.*"))
                {
                    rate = 1000000000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleGiga;
                }
                else if (Regex.IsMatch(limitUnit, "^T.*"))
                {
                    rate = 1000000000000;
                    limitUnit = RemoveScale(limitUnit);
                    limitScale = CommonConst.ScaleTera;
                }

                limitUnit = limitUnit.ToUpper();
                if (limitUnit == "HZ") limitUnit = "Hz";
                else if (limitUnit == "OHM") limitUnit = "Ohm";
                else if (limitUnit == "OHMS") limitUnit = "Ohms";
                double value;
                if (double.TryParse(limitNum, out value)) return (value * rate).ToString("G");
            }

            return limitStr;
        }

        public static string RemoveScale(string limitUnit)
        {
            return limitUnit.Substring(1, limitUnit.Length - 1);
        }

        public static string ConvertUseLimit(string limitStr, out string limitUnit, out string limitScale)
        {
            long value;
            if (ConvertNumber(limitStr, out value))
            {
                limitUnit = "";
                limitScale = "";
                return value.ToString();
            }

            return ConvertUseLimitToGlbSpec(limitStr, out limitUnit, out limitScale);
        }

        public static string ConvertUseLimitToGlbSpec(string limitStr, out string limitUnit, out string limitScale)
        {
            limitUnit = "";
            limitScale = "";
            string result;
            {
                {
                    var matches = Regex.Matches(limitStr, @"[\w|.]+");
                    result = limitStr;
                    var replaceList = new List<string>();
                    foreach (Match m in matches)
                        if (m.Value.Trim().ToUpper().StartsWith("VDD"))
                        {
                            if (!replaceList.Contains(m.Value))
                            {
                                result = result.Replace(m.Value, "_" + m.Value.ToUpper() + Var);
                                replaceList.Add(m.Value);
                            }
                        }
                        else
                        {
                            if (!replaceList.Contains(m.Value))
                            {
                                result = result.Replace(m.Value, ConvertUnits(m.Value, out limitUnit, out limitScale));
                                replaceList.Add(m.Value);
                            }
                        }

                    if (result.Contains(Var))
                        result = "=" + result;
                }
            }
            return result;
        }

        private static bool ConvertNumber(string text, out long value)
        {
            value = 0;
            if (text.Length <= 2) return false;

            var prefix = text.Substring(0, 2).ToLower();
            var number = text.Remove(0, 2);
            try
            {
                switch (prefix)
                {
                    case "0b":
                        value = Convert.ToInt64(number, 2);
                        return true;
                    case "0x":
                        value = Convert.ToInt64(number, 16);
                        return true;
                    case "0d":
                        value = Convert.ToInt64(number);
                        return true;
                }
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }

        public static string ConvertValueSpec(string value)
        {
            return value.Replace(" ", "");
        }

        public static string ConvertForceValueToGlbSpec(ForcePin forcePin)
        {
            var result = "";
            {
                var forceValue = forcePin.ForceValue;
                if (!forceValue.ToUpper().Contains("VDD") && !forceValue.ToUpper().Contains("PINS"))
                    return ConvertUnits(forceValue);
                if (forceValue.ToUpper().Contains("VDD"))
                {
                    var reg = new Regex(@"VDD\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + Var);
                }
                else if (forceValue.ToUpper().Contains("PINS"))
                {
                    var reg = new Regex(@"\w*PINS\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + Var);
                }
            }

            return result;
        }

        public static string ConvertValueWithGlbSpec(string value)
        {
            var result = "";
            {
                var forceValue = value;
                if (!forceValue.ToUpper().Contains("VDD") && !forceValue.ToUpper().Contains("PINS"))
                    return ConvertUnits(forceValue);
                if (forceValue.ToUpper().Contains("VDD"))
                {
                    var reg = new Regex(@"VDD\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + Var);
                }
                else if (forceValue.ToUpper().Contains("PINS"))
                {
                    var reg = new Regex(@"\w*PINS\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + Var);
                }
            }

            return result;
        }

        public static List<MeasPin> SortMeasPin(List<MeasPin> pinList)
        {
            var sortPins = new List<MeasPin>();
            if (pinList == null || pinList.Count == 0)
                return sortPins;
            var seqCount = pinList.Max(p => p.SequenceIndex);

            for (var i = 1; i <= seqCount; i++)
            {
                var seqPins = pinList.Where(p => p.SequenceIndex == i).ToList();
                if (seqPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff)))
                {
                    sortPins.AddRange(seqPins.Where(p => p.MeasType.Equals(MeasType.MeasV)).ToList()
                        .OrderBy(p => p.PinName));
                    sortPins.AddRange(seqPins.Where(p => p.MeasType.Equals(MeasType.MeasVdiff)).ToList()
                        .OrderBy(p => p.PinName.Split(':')[0]));
                    sortPins.AddRange(seqPins.Where(p => p.MeasType.Equals(MeasType.MeasVocm)).ToList()
                        .OrderBy(p => p.PinName.Split(':')[0]));
                }
                else
                {
                    var pins = seqPins.Select(p => p.PinName).ToList();

                    if (pinList.Where(p => p.SequenceIndex == i).Any(p => string.IsNullOrEmpty(p.CusStr)))
                        pins.Sort(delegate(string x, string y)
                        {
                            var a = x;
                            var b = y;
                            if (a.Contains("::")) //If pin name is "Pin_P::Pin_N"
                                a = a.Split(':')[0];
                            else if (a.Contains(":")) //If pin Name is "FT:VDD_FIXED"
                                a = a.Split(':')[1];
                            else if (a.Contains("=")) //If pin Name is "FT=VDD_FIXED"
                                a = a.Split('=')[1];
                            if (b.Contains("::"))
                                b = b.Split(':')[0];
                            else if (b.Contains(":"))
                                b = b.Split(':')[1];
                            else if (b.Contains("="))
                                b = b.Split('=')[1];
                            return string.CompareOrdinal(a, b);
                        });

                    foreach (var pin in pins)
                    {
                        var targetPin = pinList.FirstOrDefault(p => p.PinName.Equals(pin) && p.SequenceIndex == i);
                        var copyTargetPin = new MeasPin();
                        if (targetPin != null)
                        {
                            copyTargetPin.PinName = targetPin.PinName;
                            copyTargetPin.SequenceIndex = targetPin.SequenceIndex;
                            copyTargetPin.Copy(targetPin);
                        }

                        sortPins.Add(copyTargetPin);
                    }
                }
            }

            return sortPins;
        }

        private static void RedoJobRange(ref List<string> jobPartList, int seq, int index)
        {
            var currentRangeList = new List<string>();
            foreach (var jobMeasStr in jobPartList)
            {
                var job = jobMeasStr.Split('=')[0] + "=";
                var item = jobMeasStr.Replace(job, "");
                var seqList = item.Split('+');
                var targetSeq = seqList[seq - 1];
                var targetPinList = targetSeq.Split(',').ToList();
                targetPinList.RemoveAt(index);
                var tmp = string.Join(",", targetPinList);
                seqList[seq - 1] = tmp;
                currentRangeList.Add(job + string.Join("+", seqList));
            }

            jobPartList = currentRangeList;
        }

        public static string RedefineRange(HardIpPattern pattern, HardIpReference info, string iRangeStr)
        {
            if (string.IsNullOrEmpty(iRangeStr) || !iRangeStr.Contains("CP") && !iRangeStr.Contains("FT"))
                return iRangeStr;

            var cpPart = iRangeStr.Split(';').ToList().FindAll(p => p.Contains("CP"));
            var ftPart = iRangeStr.Split(';').ToList().FindAll(p => p.Contains("FT"));

            if (pattern.MeasPins.Exists(p => p.PinName.Contains("CP=") || p.PinName.Contains("FT=")))
            {
                var seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                if (seqCount > 0)
                    for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var measPinList = pattern.MeasPins
                            .Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC").ToList();
                        for (var i = 0; i < measPinList.Count; i++)
                            if (measPinList[i].PinName.Contains("CP=") || measPinList[i].PinName.Contains("FT="))
                            {
                                var testJob = measPinList[i].PinName.Split('=')[0];
                                if (testJob == "FT")
                                {
                                    var measCpPinList = pattern.MeasPins.Where(a =>
                                        a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC" &&
                                        !a.PinName.StartsWith("FT", StringComparison.OrdinalIgnoreCase)).ToList();
                                    if (measCpPinList.Exists(p => measPinList[i].PinName.Contains(p.PinName)))
                                        RedoJobRange(ref ftPart, sequenceIndex, i);
                                }
                            }
                    }

                var cpStr = string.Join(";", cpPart);
                var ftStr = string.Join(";", ftPart);
                var resultStr = SortCpFtCurrentRange(cpStr + ";" + ftStr);
                return resultStr;
            }

            var result = SortCpFtCurrentRange(iRangeStr);
            return result;
        }

        private static string SortCpFtCurrentRange(string fullRange)
        {
            var dicSeq = new Dictionary<int, string>();
            var jobRange = fullRange.Split(';').ToList();

            for (var i = 0; i < jobRange.Count; i++)
            {
                var job = jobRange[i].Split('=')[0] + "=";
                var seqList = jobRange[i].Split('+').ToList();
                for (var j = 0; j < seqList.Count; j++)
                    if (!seqList[j].Contains(job))
                        seqList[j] = job + seqList[j];
                jobRange[i] = string.Join("+", seqList);
            }

            for (var i = 0; i < jobRange.Count; i++)
            {
                var seqList = jobRange[i].Split('+').ToList();
                for (var j = 0; j < seqList.Count; j++)
                    if (dicSeq.ContainsKey(j))
                        dicSeq[j] += ";" + seqList[j];
                    else
                        dicSeq.Add(j, seqList[j]);
            }

            var result = dicSeq.Aggregate("", (current, seq) => current + seq.Value + "+");
            return result.TrimEnd('+');
        }

        public static string SortCpFtPin(string measStr)
        {
            var needSort = false;
            var seqList = measStr.Split(new[] {'+'}, StringSplitOptions.RemoveEmptyEntries).ToList();
            seqList.ForEach(p =>
            {
                if (p.Split(',').ToList().Exists(k => k.Contains("CP=") || k.Contains("FT="))) needSort = true;
            });

            if (!needSort) return measStr;
            var newSeqList = new List<string>();
            var measSeqList = measStr.Split('+').ToList();

            foreach (var seqPin in measSeqList)
            {
                var cpMeasList = new List<string>();
                var ftMeasList = new List<string>();
                if (string.IsNullOrWhiteSpace(seqPin))
                {
                    newSeqList.Add("");
                    continue;
                }

                if (!seqPin.Contains("CP=") && !seqPin.Contains("FT="))
                {
                    newSeqList.Add(seqPin);
                }
                else
                {
                    if (seqPin.Contains(":")) //sweep pin contain ":"
                    {
                        foreach (var pin in seqPin.Split(','))
                        foreach (var sweepPin in pin.Split(':'))
                            if (sweepPin.Contains("CP="))
                            {
                                cpMeasList.Add(sweepPin.Replace("CP=", ""));
                            }
                            else if (sweepPin.Contains("FT="))
                            {
                                ftMeasList.Add(sweepPin.Replace("FT=", ""));
                            }
                            else
                            {
                                cpMeasList.Add(sweepPin);
                                ftMeasList.Add(sweepPin);
                            }

                        newSeqList.Add(
                            "CP=" + string.Join(":", cpMeasList) + ";" + "FT=" + string.Join(":", ftMeasList));
                    }
                    else
                    {
                        foreach (var pin in seqPin.Split(','))
                            if (pin.Contains("CP="))
                            {
                                cpMeasList.Add(pin.Replace("CP=", ""));
                            }
                            else if (pin.Contains("FT="))
                            {
                                ftMeasList.Add(pin.Replace("FT=", ""));
                            }
                            else
                            {
                                cpMeasList.Add(pin);
                                ftMeasList.Add(pin);
                            }

                        var tmpCpMeasList = cpMeasList.Distinct();
                        var tmpFtMeasList = ftMeasList.Distinct();

                        newSeqList.Add("CP=" + string.Join(",", tmpCpMeasList) + ";" + "FT=" +
                                       string.Join(",", tmpFtMeasList));
                    }
                }
            }

            return string.Join("+", newSeqList);
        }

        public static void SortIdsPins(List<MeasPin> pinList)
        {
            var newPinList = new List<MeasPin>();
            var corePowers = new List<string>();
            if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey("COREPOWER"))
                corePowers = HardIpDataMain.TestPlanData.PinGroupList["COREPOWER"];
            var otherPowers = new List<string>();
            if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey("OTHERPOWER"))
                otherPowers = HardIpDataMain.TestPlanData.PinGroupList["OTHERPOWER"];
            foreach (var pinName in corePowers)
            {
                var pin = pinList.Find(a => a.PinName == pinName);
                if (pin != null)
                {
                    newPinList.Add(pin);
                    pinList.Remove(pin);
                }

                //pinName.newPinList.Add(from a in pinList where a.PinName == pinName select a);
            }

            foreach (var pinName in otherPowers)
            {
                var pin = pinList.Find(a => a.PinName == pinName);
                if (pin != null)
                {
                    newPinList.Add(pin);
                    pinList.Remove(pin);
                }

                //pinName.newPinList.Add(from a in pinList where a.PinName == pinName select a);
            }

            pinList.AddRange(newPinList);
        }

        public static string RemoveDummyPlusSign(string measPins)
        {
            var pins = Regex.Split(measPins, @"\++").ToList();
            //pins = pins.Where(s => !string.IsNullOrEmpty(s)).ToList();
            var count = pins.Distinct().ToList().Count;
            if (count == 1)
                return pins[0].Trim(',');
            return measPins.Trim(',');
        }

        public static string RemoveDummyForceV(string forceV, string regexString)
        {
            if (forceV != "")
            {
                var values = Regex.Split(forceV, regexString).ToList();
                var count = values.Distinct().ToList().Count;
                if (count == 1)
                    return values[0];
                return forceV;
            }

            return forceV;
        }

        public static string RemoveDummy(string value, string regexString)
        {
            if (value != "")
            {
                var valueList = Regex.Split(value, regexString, RegexOptions.IgnoreCase).ToList();
                if (valueList.All(x => x == ""))
                    return "";
                return value;
            }

            return value;
        }

        public static string ConvertDifferentialPinGroup(string measPins)
        {
            var result = measPins;
            if (!measPins.Contains("::") && HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(measPins))
            {
                result = "";
                var pinList = SearchInfo.DecomposeGroups(measPins);
                for (var i = 0; i < pinList.Count; i++)
                {
                    result += pinList[i + 1] + "::" + pinList[i] + ",";
                    i++;
                }

                result = result.TrimEnd(',');
            }

            return result;
        }
    }
}