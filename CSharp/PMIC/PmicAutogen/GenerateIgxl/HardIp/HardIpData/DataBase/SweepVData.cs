using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class StepEvaluator
    {
        public static string ErrMSg;

        public static int Evaluate(string inStr)
        {
            ErrMSg = "";
            var tmp = inStr.Replace(" ", "").Split(',').ToArray();
            var start = Replace(Replace(tmp[0], @"-"), @"+");
            var stop = Replace(Replace(tmp[1], @"-"), @"+");
            if (!Regex.IsMatch(start, "^-")) start = "+" + start;
            if (!Regex.IsMatch(stop, "^-")) stop = "+" + stop;
            var step = new List<string>();
            step.Add(tmp[2]);
            var starts = Regex.Split(start, @",").ToList();
            var stops = Regex.Split(stop, @",").ToList();
            var checks = starts.Select(a => a).ToList();
            foreach (var item in checks)
                if (stops.Contains(item))
                {
                    starts.Remove(item);
                    stops.Remove(item);
                }

            //TODO                
            //`es
            var sweepRange = EvaluateExpression(GetCalStr(starts, stops));
            var stepSize = EvaluateExpression(CalStr(step));
            return Convert.ToInt32(sweepRange / stepSize);
        }

        private static double EvaluateExpression(string eqn)
        {
            var dt = new DataTable();
            try
            {
                var result = dt.Compute(eqn, string.Empty);
                return Convert.ToDouble(result);
            }
            catch (Exception ex)
            {
                ErrMSg = "Error! Can't do data table EvaluateExpression! >> " + eqn + ex;
                return 0;
            }
        }

        private static string GetCalStr(List<string> l1, List<string> l2)
        {
            var evalStrSt = CalStr(l1);
            var evalStrSp = CalStr(l2);
            return evalStrSp + "-(" + evalStrSt + ")";
        }

        private static string CalStr(List<string> l1)
        {
            var evalStr = "";
            foreach (var item in l1)
            {
                var tmp = Regex.Replace(item, "V|A|Hz|OHM|S", "", RegexOptions.IgnoreCase);
                evalStr = evalStr + tmp;
            }

            return evalStr;
        }

        private static string Replace(string oldText, string newText)
        {
            return Regex.Replace(oldText, "\\" + newText, "," + newText);
        }
    }

    public class SweepVData
    {
        public SweepVData(string sweepStr)
        {
            IsEquation = false;
            //DataConvertor.ConvertForceValueToGlbSpec(
            if (sweepStr.Split(':').Length == 2)
            {
                PinName = DataConvertor.ConvertValueWithGlbSpec(sweepStr.Split(':')[0]);
                var info = sweepStr.Split(':')[1];
                if (info.Split(',').Length == 3)
                {
                    Start = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[0]);
                    Stop = DataConvertor.ConvertValueWithGlbSpec(info.Split(',')[1]);
                    Step = IsEquation
                        ? StepEvaluator.Evaluate(info).ToString()
                        : DataConvertor.ConvertValueWithGlbSpec(Operand == "-"
                            ? info.Split(',')[2].Replace("-", "")
                            : info.Split(',')[2]);
                }
                else
                {
                    Start = "0";
                    Stop = "0";
                    Step = "0";
                }
            }
        }

        public string PinName { get; set; }
        public string Type { get; set; }
        public string Start { get; set; }
        public string Stop { get; set; }
        public string Step { get; set; }
        public string Axis { get; set; }
        public bool IsEquation { get; set; }

        public string Operand
        {
            get
            {
                if (_IsValidDouble(Start) && _IsValidDouble(Stop))
                    return double.Parse(Start) < double.Parse(Stop) ? "+" : "-";
                IsEquation = true;
                return "+";
            }
            set { throw new NotImplementedException(); }
        }

        private bool _IsValidDouble(string value)
        {
            double tmp;
            return double.TryParse(value, out tmp);
        }
    }
}