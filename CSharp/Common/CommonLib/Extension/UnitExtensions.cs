using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static class UnitExtensions
    {
        private const string UnitVolt = "V";
        private const string UnitMiniVolt = "mV";
        private const string UnitMicroVolt = "uV";

        private const string UnitAmpere = "A";
        private const string UnitMiniAmpere = "mA";
        private const string UnitMicroAmpere = "uA";
        private const string UnitNanoAmpere = "nA";
        private const string UnitPicoAmpere = "pA";

        private const string UnitHz = "Hz";
        private const string UnitKhz = "KHz";
        private const string UnitMhz = "MHz";

        private const string ScaleNano = "n";
        private const string ScaleMicro = "u";
        private const string ScaleMilli = "m";

        private const string ScaleKilo = "K";
        private const string ScaleMega = "M";
        private const string ScaleGiga = "G";
        private const string ScaleTera = "T";

        private const string regPattern = @"(?<value>[+-]?\d*[.]?\d+)\s*(?<unit>\w*)\s*?";

        public static bool TryConvertToFreq(this string source, out string outputValue)
        {
            var value = Regex.Match(source, regPattern).Groups["value"].ToString();
            var unit = Regex.Match(source, regPattern).Groups["unit"].ToString();
            double number;
            outputValue = string.Empty;
            if (double.TryParse(value, out number) == false) return false;

            if (unit.Equals(UnitHz, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = source;
            }
            else if (unit.Equals(UnitKhz, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (number * 1e3).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UnitMhz, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = (number * 1e6).ToString(CultureInfo.InvariantCulture);
            }
            else if (string.IsNullOrEmpty(unit))
            {
                outputValue = value;
                return true;
            }
            else
            {
                outputValue = source;
                return false;
            }
            return true;
        }

        public static bool TryConvertToVolt(this string source, out string outputValue)
        {
            var value = Regex.Match(source, regPattern).Groups["value"].ToString();
            var unit = Regex.Match(source, regPattern).Groups["unit"].ToString();
            double number;
            outputValue = string.Empty;
            if (double.TryParse(value, out number) == false) return false;

            if (unit.Equals(UnitVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = value;
            }
            else if (unit.Equals(UnitMiniVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(UnitMicroVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else if (string.IsNullOrEmpty(unit))
            {
                outputValue = value;
                return true;
            }
            else
            {
                outputValue = source;
                return false;
            }
            return true;
        }

        public static bool TryCombineVolt(this string source, string unit, out string outputValue)
        {
            double number;
            outputValue = string.Empty;
            if (double.TryParse(source, out number) == false) return false;

            if (unit.Equals(UnitVolt, StringComparison.OrdinalIgnoreCase))
                outputValue = source;
            else if (unit.Equals(UnitMiniVolt, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(number / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(UnitMicroVolt, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(number / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            else
                return false;
            return true;
        }

        public static bool TryCombineAmpere(this string source, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            double lDValue;
            if (double.TryParse(source, out lDValue) == false) return false;

            if (unit.Equals(UnitAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = source;
            else if (unit.Equals(UnitMiniAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(UnitMicroAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(UnitNanoAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e9, 12).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(UnitPicoAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e9, 15).ToString(CultureInfo.InvariantCulture);
            else
                return false;
            return true;
        }

        public static string ConvertNumber(this string source)
        {
            if (source.Contains("10^"))
                source = source.Replace("*10^", "E");
            if (source == "" || source.Contains("E") ||
                Regex.IsMatch(source, @"^(\d|\.|-)+$")) //Limit value may be 1.2E-5
                return source;
            if (Regex.IsMatch(source, @"^(\d|\.|-)+(\w)*$"))
            {
                var number = Regex.Match(source, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (number == "0")
                    return number;
                var unit = source.Replace(number, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(unit, "^m.*"))
                    rate = 1 / (double)1000;
                else if (Regex.IsMatch(unit, "^u.*"))
                    rate = 1 / (double)1000000;
                else if (Regex.IsMatch(unit, "^n.*"))
                    rate = 1 / (double)1000000000;
                else if (Regex.IsMatch(unit.ToLower(), "^k.*"))
                    rate = 1000;
                else if (Regex.IsMatch(unit, "^M.*"))
                    rate = 1000000;
                else if (Regex.IsMatch(unit, "^G.*"))
                    rate = 1000000000;
                double value;
                if (double.TryParse(number, out value))
                    return (value * rate).ToString("G");
            }
            return source;
        }

        public static string ConvertUnit(this string source, out string unit, out string scale)
        {
            unit = "";
            scale = "";
            if (source == "" || source.Contains("E") || Regex.IsMatch(source, @"^(\d|\.|-)+$")) //Limit value may be 1.2E-5
                return source;
            if (Regex.IsMatch(source, @"^(\d|\.|-)+(\w)*$"))
            {
                var number = Regex.Match(source, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (number == "0")
                    return number;
                unit = source.Replace(number, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(unit, "^m.*"))
                {
                    rate = 1 / (double)1000;
                    unit = RemoveScale(unit);
                    scale = ScaleMilli;
                }
                else if (Regex.IsMatch(unit, "^u.*"))
                {
                    rate = 1 / (double)1000000;
                    unit = RemoveScale(unit);
                    scale = ScaleMicro;
                }
                else if (Regex.IsMatch(unit, "^n.*"))
                {
                    rate = 1 / (double)1000000000;
                    unit = RemoveScale(unit);
                    scale = ScaleNano;
                }
                else if (Regex.IsMatch(unit.ToLower(), "^k.*"))
                {
                    rate = 1000;
                    unit = RemoveScale(unit);
                    scale = ScaleKilo;
                }
                else if (Regex.IsMatch(unit, "^M.*"))
                {
                    rate = 1000000;
                    unit = RemoveScale(unit);
                    scale = ScaleMega;
                }
                else if (Regex.IsMatch(unit, "^G.*"))
                {
                    rate = 1000000000;
                    unit = RemoveScale(unit);
                    scale = ScaleGiga;
                }
                else if (Regex.IsMatch(unit, "^T.*"))
                {
                    rate = 1000000000000;
                    unit = RemoveScale(unit);
                    scale = ScaleTera;
                }

                unit = unit.ToUpper();
                if (unit == "HZ") unit = "Hz";
                else if (unit == "OHM") unit = "Ohm";
                else if (unit == "OHMS") unit = "Ohms";
                double value;
                if (double.TryParse(number, out value)) return (value * rate).ToString("G");
            }

            return source;
        }

        private static string RemoveScale(string limitUnit)
        {
            return limitUnit.Substring(1, limitUnit.Length - 1);
        }
    }
}