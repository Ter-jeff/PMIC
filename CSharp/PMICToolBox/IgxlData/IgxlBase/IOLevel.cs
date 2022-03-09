namespace IgxlData.IgxlBase
{
    public class IoLevel
    {
        #region Field
        private string _pinName;
        private string _vil;
        private string _vih;
        private string _vol;
        private string _voh;
        private string _vohAlt1;
        private string _vohAlt2;
        private string _iol;
        private string _ioh;
        private string _vt;
        private string _vcl;
        private string _vch;
        private string _voutLoTyp;
        private string _voutHiTyp;
        private string _driverMode;
        #endregion

        #region Property

        public string PinName
        {
            get { return _pinName; }
            set { _pinName = value; }
        }

        public string vil
        {
            get { return _vil; }
            set { _vil = value; }
        }

        public string vih
        {
            get { return _vih; }
            set { _vih = value; }
        }

        public string vol
        {
            get { return _vol; }
            set { _vol = value; }
        }

        public string voh
        {
            get { return _voh; }
            set { _voh = value; }
        }

        public string voh_alt1
        {
            get { return _vohAlt1; }
            set { _vohAlt1 = value; }
        }

        public string voh_alt2
        {
            get { return _vohAlt2; }
            set { _vohAlt2 = value; }
        }

        public string iol
        {
            get { return _iol; }
            set { _iol = value; }
        }

        public string ioh
        {
            get { return _ioh; }
            set { _ioh = value; }
        }

        public string vt
        {
            get { return _vt; }
            set { _vt = value; }
        }

        public string vcl
        {
            get { return _vcl; }
            set { _vcl = value; }
        }

        public string vch
        {
            get { return _vch; }
            set { _vch = value; }
        }

        public string voutLoTyp
        {
            get { return _voutLoTyp; }
            set { _voutLoTyp = value; }
        }

        public string voutHiTyp
        {
            get { return _voutHiTyp; }
            set { _voutHiTyp = value; }
        }

        public string driverMode
        {
            get { return _driverMode; }
            set { _driverMode = value; }
        }

        #endregion

        #region Constructor

        public IoLevel(string pinName, string vil, string vih, string vol, string voh, string vohAlt1, string vohAtl2, string iol,
            string ioh, string vt, string vcl, string vch, string voutLoTyp, string voutHiTyp, string driverMode)
        {
            _pinName = pinName;
            _vil = vil;
            _vih = vih;
            _vol = vol;
            _voh = voh;
            _vohAlt1 = vohAlt1;
            _vohAlt2 = vohAtl2;
            _iol = iol;
            _ioh = ioh;
            _vt = vt;
            _vcl = vcl;
            _vch = vch;
            _voutLoTyp = voutLoTyp;
            _voutHiTyp = voutHiTyp;
            _driverMode = driverMode;
        }

        #endregion
    }
}