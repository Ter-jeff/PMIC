using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace IgxlData.Others
{
    public static class CommonConst
    {
        public const string UnitVolt = "V";
        public const string UnitMiniVolt = "mV";
        public const string UnitMicroVolt = "uV";

        public const string UnitAmpere = "A";
        public const string UnitMiniAmpere = "mA";
        public const string UnitMicroAmpere = "uA";
        public const string UnitNanoAmpere = "nA";
        public const string UnitPicoAmpere = "pA";

        public const string UnitHz = "HZ";
        public const string UnitKhz = "KHZ";
        public const string UnitMhz = "MHZ";

        //public const string ScaleFemto = "f";
        //public const string ScalePico = "p";
        public const string ScaleNano = "n";
        public const string ScaleMicro = "u";
        public const string ScaleMilli = "m";

        //public const string ScalePercent = "%";
        public const string ScaleKilo = "K";
        public const string ScaleMega = "M";
        public const string ScaleGiga = "G";
        public const string ScaleTera = "T";

        //public const string PattenIsNumber = "^[0-9]+([.]{1}[0-9]+){0,1}";
        //public const string PattenIsInt = "^[0-9]+";
        //public const string PattenIsDecimal = "^[0-9]+[.][0-9]+";
        //public const string PattenGetInt = "^[^0-9]*(?<int>{[0-9]+})[^0-9]*";
        //public const string PattenGetDecimal = "^[^0-9]*(?<dec>{[0-9]+[.][0-9]+}[^0-9])*";
    }

    public class DataConvertor
    {
        public static string ConvertUnits(string limitStr)
        {
            if (limitStr.Contains("10^"))
                limitStr = limitStr.Replace("*10^", "E");
            if (limitStr == "" || limitStr.Contains("E") ||
                Regex.IsMatch(limitStr, @"^(\d|\.|-)+$")) //Limit value may be 1.2E-5
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
            var var = "_VAR";
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
                                result = result.Replace(m.Value, "_" + m.Value.ToUpper() + var);
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

                    if (result.Contains(var))
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
    }
}