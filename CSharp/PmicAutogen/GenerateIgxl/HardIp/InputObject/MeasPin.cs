using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using IgxlData.Others;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class MeasPin
    {
        #region Property

        public int RowNum { get; set; }
        public string PinName { get; set; }
        public string CusStr { get; set; }
        public string CapBit { get; set; }
        public string MeasType { get; set; }
        public int RowNumForMergeMeas { get; set; }
        public List<ForceCondition> ForceConditions { get; set; }
        public string Job { get; set; }
        public int PinCount { get; set; }

        public int RepeatCount { get; set; }

        //Added on 4/12 for new limit value format like "0.5V,0.1V,0.3V"
        public List<MeasLimit> MeasLimitsH { get; set; }
        public List<MeasLimit> MeasLimitsL { get; set; }
        public List<MeasLimit> MeasLimitsN { get; set; }
        public string LowLimit { get; set; }
        public string HighLimit { get; set; }
        public List<CurrentRange> CurrentRangeList { get; set; }
        public List<CurrentRange> CurrentRangeListH { get; set; }
        public List<CurrentRange> CurrentRangeListL { get; set; }
        public List<CurrentRange> CurrentRangeListN { get; set; }
        public int SequenceIndex { get; set; }

        public int VisitedTime { get; set; } // For sort use-limit

        //start: For forceCondition Merged 
        public string PinType { get; set; }

        //end: For forceCondition Merged 
        public string TestName { get; set; }

        //For Calc, Limit row
        public string CalcEqn { get; set; }

        //For Calc, Limit row
        public bool IsUsedPin { get; set; }

        //For FW, Interpose
        public string InterPoseFunc;
        public string RfInterPose = "";
        public string RfInstrumentSetup;
        public string MeasWaitTime;
        public string MeasRange;
        public string SkipUnit = "";
        public string MiscInfo { get; set; }

        #endregion

        #region Constructor

        public MeasPin()
        {
            PinCount = 0;
            PinName = "";
            CusStr = "";
            MeasType = "";
            Job = "";
            ForceConditions = new List<ForceCondition>();
            LowLimit = "";
            HighLimit = "";
            CurrentRangeList = new List<CurrentRange>();
            CurrentRangeListH = new List<CurrentRange>();
            CurrentRangeListL = new List<CurrentRange>();
            CurrentRangeListN = new List<CurrentRange>();
            MeasLimitsH = new List<MeasLimit>();
            MeasLimitsL = new List<MeasLimit>();
            MeasLimitsN = new List<MeasLimit>();
            VisitedTime = 1;
            SequenceIndex = 0;
            PinType = "";
            TestName = "";
            CalcEqn = "";
            IsUsedPin = false;
            MeasWaitTime = "";
            MeasRange = "";
            RfInstrumentSetup = "";
            RepeatCount = 0;
            SkipUnit = "";
            MiscInfo = "";
            InterPoseFunc = "";
        }

        public MeasPin(string pinName, string measType)
        {
            PinCount = 0;
            PinName = pinName;
            CusStr = "";
            MeasType = measType;
            Job = "";
            ForceConditions = new List<ForceCondition>();
            LowLimit = "";
            HighLimit = "";
            CurrentRangeList = new List<CurrentRange>();
            CurrentRangeListH = new List<CurrentRange>();
            CurrentRangeListL = new List<CurrentRange>();
            CurrentRangeListN = new List<CurrentRange>();
            MeasLimitsH = new List<MeasLimit>();
            MeasLimitsL = new List<MeasLimit>();
            MeasLimitsN = new List<MeasLimit>();
            VisitedTime = 1;
            SequenceIndex = 0;
            PinType = "";
            TestName = "";
            CalcEqn = "";
            MeasWaitTime = "";
            MeasRange = "";
            RfInstrumentSetup = "";
            RepeatCount = 0;
            RfInterPose = "";
            MiscInfo = "";
            InterPoseFunc = "";
        }

        public void Copy(MeasPin pin)
        {
            //PinName = pin.PinName;
            PinCount = pin.PinCount;
            MeasType = pin.MeasType;
            RowNumForMergeMeas = pin.RowNumForMergeMeas;
            CusStr = pin.CusStr;
            Job = pin.Job;
            RowNum = pin.RowNum;
            ForceConditions = pin.ForceConditions;
            LowLimit = pin.LowLimit;
            HighLimit = pin.HighLimit;
            CurrentRangeList = pin.CurrentRangeList;
            CurrentRangeListH = pin.CurrentRangeListH;
            CurrentRangeListL = pin.CurrentRangeListL;
            CurrentRangeListN = pin.CurrentRangeListN;
            MeasLimitsH = pin.MeasLimitsH;
            MeasLimitsL = pin.MeasLimitsL;
            MeasLimitsN = pin.MeasLimitsN;
            VisitedTime = pin.VisitedTime;
            PinType = pin.PinType;
            TestName = pin.TestName;
            CalcEqn = pin.CalcEqn;
            InterPoseFunc = pin.InterPoseFunc;
            RfInterPose = pin.RfInterPose;
            MeasWaitTime = pin.MeasWaitTime;
            MeasRange = "";
            RfInstrumentSetup = pin.RfInstrumentSetup;
            RepeatCount = pin.RepeatCount;
            MiscInfo = pin.MiscInfo;
        }

        public MeasPin DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as MeasPin;
            }
        }

        #endregion

        #region Set currentRange

        public List<CurrentRange> GetCurrentRangeListByVoltage(List<MeasLimit> measLimits)
        {
            var forcePinList = GetForcePinList();
            var forceValueList = GetForceValueList(forcePinList);

            var currentList = new List<CurrentRange>();
            if (MeasType.Equals("measi", StringComparison.OrdinalIgnoreCase) ||
                MeasType.Equals("measidiff", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(MeasType, @"MeasR[1|2]", RegexOptions.IgnoreCase))
            {
                for (var i = 0; i < MeasLimitsH.Count; i++)
                {
                    var highLimitList = new List<string>();
                    var lowLimitList = new List<string>();

                    GetLimitList(measLimits[i], ref lowLimitList, ref highLimitList);

                    var hiRange = ParseLimitList(highLimitList);
                    var loRange = ParseLimitList(lowLimitList);

                    if (!(hiRange > 0 || loRange > 0)) continue;

                    if (!Regex.IsMatch(MeasType, @"^MeasR\d"))
                    {
                        var range = new CurrentRange
                        {
                            JobName = MeasLimitsL[i].JobName,
                            Value = (hiRange > loRange ? hiRange : loRange).ToString("G")
                        };
                        currentList.Add(range);
                    }
                    else
                    {
                        if (ForceConditions.Count == 0) return null;
                        if (forceValueList.Count == 0 || forceValueList.Max() == 0) return null;

                        #region if forceType == "I", currentRange = forceI value

                        if (forcePinList.All(x => x.ForceType.Equals("I", StringComparison.OrdinalIgnoreCase)))
                        {
                            var range = new CurrentRange
                            {
                                JobName = MeasLimitsL[i].JobName, Value = Math.Abs(forceValueList.Max()).ToString("G")
                            };
                            currentList.Add(range);
                        }

                        #endregion

                        #region if forceType == "V", currentRange = (forceV value)/lowest(limit range)

                        else if (forcePinList.All(x => x.ForceType.Equals("V", StringComparison.OrdinalIgnoreCase)))
                        {
                            double value;
                            if (hiRange > 0 && loRange > 0)
                                value = Math.Abs(hiRange > loRange
                                    ? loRange
                                    : hiRange); // get max irange with max voltage and min R
                            else
                                value = Math.Abs(hiRange > loRange ? hiRange : loRange); //get value that is not zero
                            var range = new CurrentRange
                            {
                                JobName = MeasLimitsL[i].JobName,
                                Value = Math.Round(forceValueList.Max() / value, 7).ToString("G")
                            };
                            currentList.Add(range);
                        }

                        #endregion
                    }
                }

                return currentList;
            }

            return null;
        }

        public List<CurrentRange> GetCurrentRangeList()
        {
            var forcePinList = GetForcePinList();
            var forceValueList = GetForceValueList(forcePinList);

            var currentList = new List<CurrentRange>();
            if (MeasType.Equals("measi", StringComparison.OrdinalIgnoreCase) ||
                MeasType.Equals("measidiff", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(MeasType, @"MeasR[1|2]", RegexOptions.IgnoreCase))
            {
                for (var i = 0; i < MeasLimitsH.Count; i++)
                {
                    var highLimitList = new List<string>();
                    var lowLimitList = new List<string>();

                    GetLimitList(MeasLimitsH[i], ref lowLimitList, ref highLimitList);
                    GetLimitList(MeasLimitsL[i], ref lowLimitList, ref highLimitList);
                    GetLimitList(MeasLimitsN[i], ref lowLimitList, ref highLimitList);

                    var hiRange = ParseLimitList(highLimitList);
                    var loRange = ParseLimitList(lowLimitList);

                    if (!(hiRange > 0 || loRange > 0)) continue;

                    if (!Regex.IsMatch(MeasType, @"^MeasR\d"))
                    {
                        var range = new CurrentRange
                        {
                            JobName = MeasLimitsL[i].JobName,
                            Value = (hiRange > loRange ? hiRange : loRange).ToString("G")
                        };
                        currentList.Add(range);
                    }
                    else
                    {
                        if (ForceConditions.Count == 0) return null;
                        if (forceValueList.Count == 0 || forceValueList.Max() == 0) return null;

                        #region if forceType == "I", currentRange = forceI value

                        if (forcePinList.All(x => x.ForceType.Equals("I", StringComparison.OrdinalIgnoreCase)))
                        {
                            var range = new CurrentRange
                            {
                                JobName = MeasLimitsL[i].JobName, Value = Math.Abs(forceValueList.Max()).ToString("G")
                            };
                            currentList.Add(range);
                        }

                        #endregion

                        #region if forceType == "V", currentRange = (forceV value)/lowest(limit range)

                        else if (forcePinList.All(x => x.ForceType.Equals("V", StringComparison.OrdinalIgnoreCase)))
                        {
                            double value;
                            if (hiRange > 0 && loRange > 0)
                                value = Math.Abs(hiRange > loRange
                                    ? loRange
                                    : hiRange); // get max irange with max voltage and min R
                            else
                                value = Math.Abs(hiRange > loRange ? hiRange : loRange); //get value that is not zero
                            var range = new CurrentRange
                            {
                                JobName = MeasLimitsL[i].JobName,
                                Value = Math.Round(forceValueList.Max() / value, 7).ToString("G")
                            };
                            currentList.Add(range);
                        }

                        #endregion
                    }
                }

                return currentList;
            }

            return null;
        }

        private List<ForcePin> GetForcePinList()
        {
            var forcePinList = new List<ForcePin>();
            foreach (var forceCondition in ForceConditions)
            foreach (var forcePin in forceCondition.ForcePins)
                if (forcePin.PinName.Equals(PinName, StringComparison.OrdinalIgnoreCase))
                {
                    forcePinList.Add(forcePin);
                }
                else
                {
                    //DecomposeGroups for force condition and Misc info match
                    var forceNameList = new List<string>();
                    forceNameList.AddRange(SearchInfo.DecomposeGroups(forcePin.PinName));

                    var newForcePinList = new List<ForcePin>();
                    foreach (var a in forceNameList)
                    {
                        var tempForcePin = forcePin.DeepClone();
                        tempForcePin.PinName = a;
                        newForcePinList.Add(tempForcePin);
                    }

                    var newForcePin =
                        newForcePinList.Find(s => s.PinName.Equals(PinName, StringComparison.OrdinalIgnoreCase));
                    if (newForcePin != null)
                        forcePinList.Add(newForcePin);
                }

            return forcePinList;
        }

        private List<double> GetForceValueList(List<ForcePin> forcePinList)
        {
            var forceValueList = new List<double>();
            foreach (var forcePin in forcePinList)
            {
                double forceValue;
                if (double.TryParse(forcePin.ForceValue, out forceValue))
                    forceValueList.Add(forceValue);
                else
                    return forceValueList;
            }

            return forceValueList;
        }

        private string ConvertToNum(string limit)
        {
            //if(!limit.Contains("*")) return limit;

            var regexPattern =
                @"(?<Value1>[+-]?\d+([.])?(\d+)?)[\+\-\*\\]?(?<Value2>[+-]?(\d+)?[.]?(\d+)?)(?<Unit>\w+)*";
            var value1 = Regex.Match(limit, regexPattern).Groups["Value1"].ToString();
            var value2 = Regex.Match(limit, regexPattern).Groups["Value2"].ToString();
            var unit = Regex.Match(limit, regexPattern).Groups["Unit"].ToString();

            double tmpValue1, tmpValue2;
            if (!double.TryParse(value1, out tmpValue1))
                return limit;

            if (!double.TryParse(value2, out tmpValue2))
                return limit;

            var value = tmpValue1 * tmpValue2;

            return value + unit;
        }

        private void GetLimitList(MeasLimit limit, ref List<string> loValueList, ref List<string> hiValueList)
        {
            var lowLimit = DataConvertor.ConvertUnits(ConvertToNum(limit.LoLimit));
            if (Regex.IsMatch(lowLimit, @"^(\d|\.|-)$") || lowLimit.Contains("E-") ||
                Regex.IsMatch(lowLimit, @"^(\d|\.|-)+$"))
                loValueList.Add(lowLimit);
            var highLimit = DataConvertor.ConvertUnits(ConvertToNum(limit.HiLimit));
            if (Regex.IsMatch(highLimit, @"^(\d|\.|-)$") || highLimit.Contains("E-") ||
                Regex.IsMatch(highLimit, @"^(\d|\.|-)+$"))
                hiValueList.Add(highLimit);
        }

        private double ParseLimitList(List<string> limitList)
        {
            double range = 0;
            if (limitList.Count > 0)
            {
                var newValueList = limitList.Select(Convert.ToDouble).ToList();
                newValueList = newValueList.Select(Math.Abs).ToList();
                newValueList.Sort();
                range = MeasType.ToUpper().Contains("R") ? newValueList[0] : newValueList[newValueList.Count - 1];
            }

            return range;
        }

        public string GetCurrentRangeByVoltage(string voltage)
        {
            List<CurrentRange> currentRangeList;
            switch (voltage)
            {
                case "HV":
                    currentRangeList = CurrentRangeListH;
                    break;
                case "LV":
                    currentRangeList = CurrentRangeListL;
                    break;
                case "NV":
                    currentRangeList = CurrentRangeListN;
                    break;
                default:
                    currentRangeList = CurrentRangeList;
                    break;
            }

            if (currentRangeList == null || currentRangeList.Count == 0) return "";

            if (currentRangeList.Select(x => x.Value).Distinct().Count() == 1)
                return currentRangeList[0].Value;

            return string.Join(";", currentRangeList.Select(x => x.JobName + ":" + x.Value));
        }

        public string GetCurrentRange()
        {
            if (CurrentRangeList == null || CurrentRangeList.Count == 0) return "";

            if (CurrentRangeList.Select(x => x.Value).Distinct().Count() == 1)
                return CurrentRangeList[0].Value;

            return string.Join(";", CurrentRangeList.Select(x => x.JobName + ":" + x.Value));
        }

        #endregion
    }
}