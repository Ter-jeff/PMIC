using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using PmicAutogen.Local;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTestConti
{
    public enum ContiType
    {
        PowerShort,
        OpenShort
    }

    public class DcTestContiRow
    {
        public List<DcTestContiSheetLimit> LimitsPmic
        {
            set { _dcTestContiSheetLimits = value; }
            get { return _dcTestContiSheetLimits ?? (_dcTestContiSheetLimits = new List<DcTestContiSheetLimit>()); }
        }

        public ContiType TestType
        {
            get
            {
                if (Regex.IsMatch(Category, Continuity, RegexOptions.IgnoreCase))
                    return ContiType.OpenShort;

                if (Regex.IsMatch(Category, PowerShort, RegexOptions.IgnoreCase))
                    return ContiType.PowerShort;

                return ContiType.PowerShort;
            }
            set { throw new NotImplementedException(); }
        }

        public bool GetForceCondition(out string result)
        {
            string conditionValue;
            string unit;
            var neg = false;
            string condition;
            result = "";
            var type = TestType;

            //Get Condition
            if (Regex.IsMatch(Condition, OpenShortSource, RegexOptions.IgnoreCase))
            {
                //ISource
                conditionValue = Regex.Match(Condition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Value]
                    .ToString();
                unit = Regex.Match(Condition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Unit].ToString();
            }
            else if (Regex.IsMatch(Condition, OpenShortSink, RegexOptions.IgnoreCase))
            {
                //ISink
                conditionValue = Regex.Match(Condition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Value]
                    .ToString();
                unit = Regex.Match(Condition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                neg = true;
            }
            else if (Regex.IsMatch(Condition, PowerShortVForce, RegexOptions.IgnoreCase))
            {
                //VForce
                conditionValue = Regex.Match(Condition, PowerShortVForce, RegexOptions.IgnoreCase).Groups[Value]
                    .ToString();
                unit = Regex.Match(Condition, PowerShortVForce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
            }
            else
            {
                return false;
            }


            //Convert Unit
            if (unit != "" && type == ContiType.OpenShort)
                //Current
                TryToConvertToAmpere(conditionValue, unit, out condition);
            else if (unit != "" && type == ContiType.PowerShort)
                //Volt
                TryToConvertToVolt(conditionValue, unit, out condition);
            else
                condition = conditionValue;

            //Positive or negative
            if (neg)
                result = "-" + condition;
            else
                result = condition;
            return true;
        }

        public bool TryToConvertToVolt(string value, string unit, out string outputValue)
        {
            double lDValue;
            outputValue = string.Empty;
            if (double.TryParse(value, out lDValue) == false) return false;

            if (unit.Equals(CommonConst.UnitVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = value;
            }
            else if (unit.Equals(CommonConst.UnitMiniVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            }
            else if (unit.Equals(CommonConst.UnitMicroVolt, StringComparison.OrdinalIgnoreCase))
            {
                outputValue = Math.Round(lDValue / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                outputValue = value;
                return false;
            }

            return true;
        }

        public bool TryToConvertToAmpere(string value, string unit, out string outputValue)
        {
            outputValue = string.Empty;
            double lDValue;
            if (double.TryParse(value, out lDValue) == false) return false;

            if (unit.Equals(CommonConst.UnitAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = value;
            else if (unit.Equals(CommonConst.UnitMiniAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e3, 6).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(CommonConst.UnitMicroAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e6, 9).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(CommonConst.UnitNanoAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e9, 12).ToString(CultureInfo.InvariantCulture);
            else if (unit.Equals(CommonConst.UnitPicoAmpere, StringComparison.OrdinalIgnoreCase))
                outputValue = Math.Round(lDValue / 1e9, 15).ToString(CultureInfo.InvariantCulture);
            else
                return false;
            return true;
        }

        public bool GetTestLimit(out string hiLimitShort, out string lowLimitShort, out string hiLimitOpen,
            out string lowLimitOpen)
        {
            hiLimitShort = "";
            lowLimitShort = "";
            hiLimitOpen = "";
            lowLimitOpen = "";
            var regPattern = @"(?<value>[+-]?\d*[.]?\d+)\s*(?<unit>\w*)\s*?";

            var limitValue = _dcTestContiSheetLimits;
            foreach (var item in limitValue)
                if (TestType == ContiType.OpenShort)
                {
                    //Open Short
                    string hiValue;
                    string hiUnit;
                    if (!item.OpenHiLimitValue.Equals(""))
                    {
                        hiValue = Regex.Match(item.OpenHiLimitValue, regPattern).Groups["value"].ToString();
                        hiUnit = Regex.Match(item.OpenHiLimitValue, regPattern).Groups["unit"].ToString();
                        TryToConvertToVolt(hiValue, hiUnit, out hiLimitOpen);
                    }

                    string loValue;
                    string loUnit;
                    if (!item.OpenLoLimitValue.Equals(""))
                    {
                        loValue = Regex.Match(item.OpenLoLimitValue, regPattern).Groups["value"].ToString();
                        loUnit = Regex.Match(item.OpenLoLimitValue, regPattern).Groups["unit"].ToString();
                        TryToConvertToVolt(loValue, loUnit, out lowLimitOpen);
                    }

                    if (!item.ShortHiLimitValue.Equals(""))
                    {
                        hiValue = Regex.Match(item.ShortHiLimitValue, regPattern).Groups["value"].ToString();
                        hiUnit = Regex.Match(item.ShortHiLimitValue, regPattern).Groups["unit"].ToString();
                        TryToConvertToVolt(hiValue, hiUnit, out hiLimitShort);
                    }

                    if (!item.ShortLoLimitValue.Equals(""))
                    {
                        loValue = Regex.Match(item.ShortLoLimitValue, regPattern).Groups["value"].ToString();
                        loUnit = Regex.Match(item.ShortLoLimitValue, regPattern).Groups["unit"].ToString();
                        TryToConvertToVolt(loValue, loUnit, out lowLimitShort);
                    }
                }

            return true;
        }

        #region Const variable

        public const string ConHeaderCategory = "Category";
        public const string ConHeaderPinGroup = "Pin Group";
        public const string ConHeaderTimeSet = "TimeSet";
        public const string ConHeaderCondition = "Condition";
        public const string ConHeaderLimit = "Limit";
        public const string ConHeaderHiOpen = "HiLimit_Open";
        public const string ConHeaderLoOpen = "LoLimit_Open";
        public const string ConHeaderHiShort = "HiLimit_Short";
        public const string ConHeaderLoShort = "LoLimit_Short";

        private const string Value = "Value";
        private const string Unit = "Unit";

        private const string OpenShortSource =
            @"Isource\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";

        private const string OpenShortSink = @"Isink\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";

        private const string PowerShortVForce =
            @"Vforce\s*=\s*(?<" + Value + @">[+-]?(\d+[.])?\d+)(?<" + Unit + @">\w+)";

        private const string Continuity = "Continuity";
        private const string PowerShort = @"Power\s*Short";

        #endregion

        #region Filed

        private List<DcTestContiSheetLimit> _dcTestContiSheetLimits = new List<DcTestContiSheetLimit>();

        public DcTestContiRow()
        {
            ColumnIdx = 0;
            RowNum = 0;
        }

        #endregion

        #region Property

        public int RowNum { get; set; }
        public int ColumnIdx { get; set; }
        public string Category { set; get; } = "";
        public string PinGroup { set; get; } = "";

        public string TimeSet { set; get; } = "";

        public string Condition { set; get; } = "";

        #endregion
    }
}