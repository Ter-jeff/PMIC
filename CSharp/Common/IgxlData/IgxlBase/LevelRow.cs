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
            _pinName = pinName;
            _parameter = parameter;
            _value = value;
            _comment = comment;
        }

        public bool IsBlankRow()
        {
            if (_comment == "" &&
                _parameter == "" &&
                _pinName == "" &&
                _value == "")
                return true;
            return false;
        }

        private string _pinName;
        private string _parameter;
        private string _value;
        private string _comment;

        public string PinName
        {
            get { return _pinName; }
            set { _pinName = value; }
        }

        public string Parameter
        {
            get { return _parameter; }
            set { _parameter = value; }
        }

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public string Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }

        public string SpecialComment { get; set; }
    }
}