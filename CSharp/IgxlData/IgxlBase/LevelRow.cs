using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class LevelRow
    {
        #region Field
        private string _pinName;
        private string _parameter;
        private string _value;
        private string _comment;
        #endregion

        #region Property

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

        #endregion

        #region Constructor

        public LevelRow(string pinName, string parameter, string value, string comment)
        {
            _pinName = pinName;
            _parameter = parameter;
            _value = value;
            _comment = comment;
        }

        #endregion

        #region Member Function

        public bool IsBlankRow()
        {
            if (_comment == "" &&
                _parameter == "" &&
                _pinName == "" &&
                _value == "")
            {
                return true;
            }
            return false;
        }

        #endregion
    }
}