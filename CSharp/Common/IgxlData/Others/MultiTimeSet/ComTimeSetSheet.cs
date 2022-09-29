using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.Others.MultiTimeSet
{
    public struct TsetEqnVarMap
    {
        public string TsetName;
        public Dictionary<string, double> DictVariable;
    }

    [Serializable]
    public class ComTimeSetBasicSheet : TimeSetBasicSheet
    {
        private List<string> _shiftInTSet = new List<string>(); // to record shift time set tset name

        public ComTimeSetBasicSheet(string sheetName)
            : base(sheetName)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode)
            : base(sheetName, timingMode)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode, string masterTimeSet, string timeDomain,
            string strobeRefSetup) :
            base(sheetName, timingMode, masterTimeSet, timeDomain, strobeRefSetup)
        {
        }

        public bool IsMultiShiftInTSet
        {
            get { return _shiftInTSet.Count > 1; }
            set { throw new NotImplementedException(); }
        }

        public string GetMultiShiftInStr
        {
            get { return string.Join(",", _shiftInTSet); }

            set { throw new NotImplementedException(); }
        }

        public List<TsetEqnVarMap> AllTsetEqnVariable
        {
            get
            {
                var allTsetEqnVariable = new List<TsetEqnVarMap>();
                var mainCommentVariable = new Dictionary<string, double>();
                foreach (var tset in Tsets)
                {
                    var comTsb = tset as ComTimeSetBasic;
                    if (comTsb == null)
                        throw new Exception("");

                    foreach (var pair in comTsb.SubCommentVariable)
                        if (!mainCommentVariable.ContainsKey(pair.Key))
                            mainCommentVariable.Add(pair.Key, pair.Value);

                    TsetEqnVarMap eqnVarMapObj;
                    eqnVarMapObj.TsetName = comTsb.Name;
                    eqnVarMapObj.DictVariable = mainCommentVariable;
                    allTsetEqnVariable.Add(eqnVarMapObj);
                }

                return allTsetEqnVariable;
            }
            set { throw new NotImplementedException(); }
        }

        public void AddShiftInTSetName(string tSet)
        {
            _shiftInTSet.Add(tSet);
        }

        public void InsertAlarmDataInFirstRow(string alarmString)
        {
            var alarmTSet = new ComTimeSetBasic();
            alarmTSet.Name = alarmString + " Please check it";
            alarmTSet.AddTimingRow(new TimingRow());
            alarmTSet.AddTimingRow(new TimingRow());
            TimeSetsData.Insert(0, alarmTSet);
        }

        public double GetMaxFrequency()
        {
            double maxFrequency = 0;
            foreach (var tset in Tsets)
            {
                var currentFrequency = GetFrequency(tset);
                if (currentFrequency > maxFrequency)
                    maxFrequency = currentFrequency;
            }

            return maxFrequency;
        }

        private double GetFrequency(TSet tset)
        {
            double period;
            if (double.TryParse(tset.CyclePeriod, out period)) return 1 / period;

            return GetFrequencyValue(tset);
        }

        private double GetFrequencyValue(TSet tset)
        {
            var equation = tset.CyclePeriod.Replace("=", "").Replace("(", "").Replace(")", "");
            var varName = Regex.IsMatch(equation, @"/_")
                ? equation.Substring(equation.IndexOf("/_", StringComparison.Ordinal) + 2)
                : equation.Substring(equation.IndexOf('/') + 1);
            if (AllTsetEqnVariable.Exists(x => x.TsetName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase)))
            {
                var tsetEqnVarMap = AllTsetEqnVariable.Find(x =>
                    x.TsetName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase));
                if (tsetEqnVarMap.DictVariable.ContainsKey(varName))
                {
                    var varDefinition = tsetEqnVarMap.DictVariable[varName].ToString(CultureInfo.InvariantCulture);
                    var digitName = Regex.IsMatch(equation, @"/_")
                        ? equation.Replace("_" + varName, varDefinition)
                        : equation.Replace(varName, varDefinition);

                    if (Regex.IsMatch(digitName, @"/"))
                    {
                        var numerator = digitName.Substring(0, digitName.IndexOf("/", StringComparison.Ordinal));
                        var denominator = digitName.Substring(digitName.IndexOf("/", StringComparison.Ordinal) + 1);
                        return Convert.ToDouble(denominator) / Convert.ToDouble(numerator);
                    }

                    return 1 / Convert.ToDouble(digitName);
                }
            }

            return 0;
        }

        public Dictionary<string, double> GeTsetData(InstanceRow instanceRow = null, SpecFinder specFinder = null)
        {
            var dic = new Dictionary<string, double>();
            foreach (var tset in TimeSetsData)
            {
                var timeSetName = tset.Name.Trim();
                var periodStr = tset.CyclePeriod;
                double period = 0;

                if (Regex.IsMatch(periodStr, @"^\=\(?[0-9]*\.?[0-9]+\/.+"))
                {
                    var frequencyName = periodStr.Split('/').Last();

                    if (Regex.IsMatch(frequencyName, @"\)$"))
                        frequencyName = Regex.Replace(frequencyName, @"\)$", "");

                    double freVal = 0;

                    if (instanceRow != null && specFinder != null)
                    {
                        double value;
                        double.TryParse(specFinder.GetValue(instanceRow, periodStr, frequencyName), out value);
                        period = value * 1000000000; //ns
                    }
                    else
                    {
                        #region by timeSet

                        foreach (var t in AllTsetEqnVariable)
                            if (timeSetName.Equals(t.TsetName.Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (var item in t.DictVariable)
                                    if (item.Key.Contains(frequencyName))
                                    {
                                        freVal = item.Value;
                                        break;
                                    }

                                break;
                            }

                        #endregion
                    }

                    if (freVal != 0)
                    {
                        var numbers = periodStr.Split('(').Last();
                        numbers = numbers.Split('/').First();
                        var num = float.Parse(numbers);
                        period = num / freVal * 1000000000; //ns
                    }
                }
                else
                {
                    period = Convert.ToDouble(periodStr);
                    period = period * 1000000000; //ns
                }

                dic.Add(timeSetName, period);
            }

            return dic;
        }

        public double GetExpectedTime(Dictionary<string, int> dic)
        {
            var timeSetData = GeTsetData();
            double expectedTime = 0;
            foreach (var item in dic)
            {
                double periodValue = 0;
                foreach (var tset in timeSetData)
                    if (item.Key.Equals(tset.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        periodValue = tset.Value;
                        break;
                    }

                expectedTime += item.Value * periodValue;
            }

            return expectedTime;
        }

        public Dictionary<string, Tuple<int, double>> GetTsetDic(Dictionary<string, int> dic,
            InstanceRow instanceRow = null, SpecFinder specFinder = null)
        {
            var tsetDic = new Dictionary<string, Tuple<int, double>>();
            var timeSetData = GeTsetData(instanceRow, specFinder);
            foreach (var item in dic)
            {
                double periodValue = 0;
                foreach (var tset in timeSetData)
                    if (item.Key.Equals(tset.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        periodValue = tset.Value;
                        break;
                    }

                if (!tsetDic.ContainsKey(item.Key))
                    tsetDic.Add(item.Key, new Tuple<int, double>(item.Value, periodValue));
            }

            return tsetDic;
        }
    }
}