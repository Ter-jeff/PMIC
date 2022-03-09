using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;

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
        public bool IsMulitShiftInTSet
        {
            get
            {
                return _shiftInTSet.Count > 1;
            }
        }

        public string GetMultiShiftInStr
        {
            get { return string.Join(",", _shiftInTSet); }

        }

        public List<TsetEqnVarMap> AllTsetEqnVariable
        {
            get
            {
                List<TsetEqnVarMap> allTsetEqnVariable = new List<TsetEqnVarMap>();
                Dictionary<string, double> mainCommentVariable = new Dictionary<string, double>();
                foreach (Tset tset in Tsets)
                {
                    ComTimeSetBasic comTsb = tset as ComTimeSetBasic;
                    if (comTsb == null)
                        throw new Exception("");

                    foreach (KeyValuePair<string, double> pair in comTsb.SubCommentVariable)
                    {
                        if (!mainCommentVariable.ContainsKey(pair.Key))
                            mainCommentVariable.Add(pair.Key, pair.Value);
                    }

                    TsetEqnVarMap eqnVarMapObj;
                    eqnVarMapObj.TsetName = comTsb.Name;
                    eqnVarMapObj.DictVariable = mainCommentVariable;
                    allTsetEqnVariable.Add(eqnVarMapObj);
                }
                return allTsetEqnVariable;
            }
        }

        public ComTimeSetBasicSheet(string sheetName)
            : base(sheetName)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode)
            : base(sheetName, timingMode)
        {
        }

        public ComTimeSetBasicSheet(string sheetName, string timingMode, string masterTimeSet, string timeDomain, string strobeRefSetup) :
            base(sheetName, timingMode, masterTimeSet, timeDomain, strobeRefSetup)
        {
        }

        private List<string> _shiftInTSet = new List<string>();  // to record shift time set tset name

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

        private double GetFrequency(Tset tset)
        {
            double period;
            if (double.TryParse(tset.CyclePeriod, out period)) return 1 / period;

            return GetFrequencyValue(tset);
        }

        private double GetFrequencyValue(Tset tset)
        {
            string equation = tset.CyclePeriod.Replace("=", "").Replace("(", "").Replace(")", "");
            string varname = "";
            if (Regex.IsMatch(equation, @"/_")) varname = equation.Substring(equation.IndexOf("/_", StringComparison.Ordinal) + 2);
            else varname = equation.Substring(equation.IndexOf('/') + 1);
            if (AllTsetEqnVariable.Exists(x => x.TsetName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase)))
            {
                var tsetEqnVarMap = AllTsetEqnVariable.Find(x => x.TsetName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase));
                if (tsetEqnVarMap.DictVariable.ContainsKey(varname))
                {
                    var varDefinition = tsetEqnVarMap.DictVariable[varname].ToString();
                    string digitname = Regex.IsMatch(equation, @"/_") ?
                        equation.Replace("_" + varname, varDefinition) :
                        equation.Replace(varname, varDefinition);

                    if (Regex.IsMatch(digitname, @"/"))
                    {
                        string numerator = digitname.Substring(0, digitname.IndexOf("/"));
                        string denominator = digitname.Substring(digitname.IndexOf("/") + 1);
                        return Convert.ToDouble(denominator) / Convert.ToDouble(numerator);
                    }
                    return 1 / Convert.ToDouble(digitname);
                }
            }
            return 0;
        }

        public Dictionary<string, double> GeTsetData(InstanceRow instanceRow = null, SpecFinder specFinder = null)
        {
            var dic = new Dictionary<string, double>();
            foreach (var tset in TimeSetsData)
            {
                string timeSetName = tset.Name.Trim();
                string periodStr = tset.CyclePeriod;
                double period = 0;

                if (Regex.IsMatch(periodStr, @"^\=\(?[0-9]*\.?[0-9]+\/.+"))//fomula
                {
                    string frequencyName = periodStr.Split('/').Last();
                   
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
                        {
                            if (timeSetName.Equals(t.TsetName.Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (var item in t.DictVariable)
                                {
                                    if (item.Key.Contains(frequencyName))
                                    {
                                        freVal = item.Value;
                                        break;
                                    }
                                }
                                break;
                            }
                        }

                        #endregion
                    }

                    if (freVal != 0)
                    {
                        string numsplit = periodStr.Split('(').Last();
                        numsplit = numsplit.Split('/').First();
                        float num = float.Parse(numsplit);
                        period = (num / freVal) * 1000000000; //ns
                    }
                }
                else
                {
                    period = Convert.ToDouble(periodStr);
                    period = period * 1000000000;//ns
                }
                dic.Add(timeSetName, period);
            }
            return dic;
        }

        public double GetExpectedTime(Dictionary<string, int> dics)
        {
            Dictionary<string, double> timesetData = GeTsetData();
            double expectedTime = 0;
            foreach (var dic in dics)
            {
                double periodValue = 0;
                foreach (var tset in timesetData)
                {
                    if (dic.Key.Equals(tset.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        periodValue = tset.Value;
                        break;
                    }
                }
                expectedTime += dic.Value * periodValue;
            }
            return expectedTime;
        }

        public  Dictionary<string, Tuple<int, double>> GetTsetDic(Dictionary<string, int> dics, InstanceRow instanceRow = null, SpecFinder specFinder = null)
        {
            var tsetDic = new Dictionary<string, Tuple<int, double>>();
            Dictionary<string, double> timesetData = GeTsetData(instanceRow, specFinder);
            foreach (var dic in dics)
            {
                double periodValue = 0;
                foreach (var tset in timesetData)
                {
                    if (dic.Key.Equals(tset.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        periodValue = tset.Value;
                        break;
                    }
                }
                if (!tsetDic.ContainsKey(dic.Key))
                    tsetDic.Add(dic.Key, new Tuple<int, double>(dic.Value, periodValue));
            }
            return tsetDic;
        }
    }
}
