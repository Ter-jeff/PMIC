using System;
using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    [Serializable]
    public class LevelRow : IgxlRow
    {

        public LevelRow(string pinName, string parameter, string value, string comment)
        {
            PinName = pinName;
            Parameter = parameter;
            Value = value;
            Comment = comment;
        }

        public bool IsBlankRow()
        {
            if (Comment == "" &&
                Parameter == "" &&
                PinName == "" &&
                Value == "")
                return true;
            return false;
        }

        public string PinName { get; set; }
        public string Parameter { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }
        public string SpecialComment { get; set; }
    }
}