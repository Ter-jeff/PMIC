using CommonLib.Extension;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
                conditionValue = Regex.Match(Condition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Value].ToString();
                unit = Regex.Match(Condition, OpenShortSource, RegexOptions.IgnoreCase).Groups[Unit].ToString();
            }
            else if (Regex.IsMatch(Condition, OpenShortSink, RegexOptions.IgnoreCase))
            {
                //ISink
                conditionValue = Regex.Match(Condition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Value].ToString();
                unit = Regex.Match(Condition, OpenShortSink, RegexOptions.IgnoreCase).Groups[Unit].ToString();
                neg = true;
            }
            else if (Regex.IsMatch(Condition, PowerShortVForce, RegexOptions.IgnoreCase))
            {
                //VForce
                conditionValue = Regex.Match(Condition, PowerShortVForce, RegexOptions.IgnoreCase).Groups[Value].ToString();
                unit = Regex.Match(Condition, PowerShortVForce, RegexOptions.IgnoreCase).Groups[Unit].ToString();
            }
            else
            {
                return false;
            }


            //Convert Unit
            if (unit != "" && type == ContiType.OpenShort)
                //Current
                conditionValue.TryCombineAmpere(unit, out condition);
            else if (unit != "" && type == ContiType.PowerShort)
                //Volt
                conditionValue.TryCombineVolt(unit, out condition);
            else
                condition = conditionValue;

            //Positive or negative
            if (neg)
                result = "-" + condition;
            else
                result = condition;
            return true;
        }




        public bool GetTestLimit(out string hiLimitShort, out string lowLimitShort, out string hiLimitOpen,
            out string lowLimitOpen)
        {
            hiLimitShort = "";
            lowLimitShort = "";
            hiLimitOpen = "";
            lowLimitOpen = "";

            var limitValue = _dcTestContiSheetLimits;
            foreach (var item in limitValue)
                if (TestType == ContiType.OpenShort)
                {
                    //Open Short
                    if (!item.OpenHiLimitValue.Equals(""))
                    {
                        item.OpenHiLimitValue.TryConvertToVolt(out hiLimitOpen);
                    }
                    if (!item.OpenLoLimitValue.Equals(""))
                    {
                        item.OpenLoLimitValue.TryConvertToVolt(out lowLimitOpen);
                    }

                    if (!item.ShortHiLimitValue.Equals(""))
                    {
                        item.ShortHiLimitValue.TryConvertToVolt(out hiLimitShort);
                    }

                    if (!item.ShortLoLimitValue.Equals(""))
                    {
                        item.ShortLoLimitValue.TryConvertToVolt(out lowLimitShort);
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

        private const string OpenShortSource = @"Isource\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";

        private const string OpenShortSink = @"Isink\s*=\s*[+-]?(?<" + Value + @">(\d+[.])?\d+)(?<" + Unit + @">\w+)";

        private const string PowerShortVForce = @"Vforce\s*=\s*(?<" + Value + @">[+-]?(\d+[.])?\d+)(?<" + Unit + @">\w+)";

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