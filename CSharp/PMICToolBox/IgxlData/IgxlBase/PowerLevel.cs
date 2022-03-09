namespace IgxlData.IgxlBase
{
    public class PowerLevel
    {
        #region Field

        private string _pinName;
        private string _vmain;
        private string _valt;
        private string _iFoldLevel;
        private string _tdelay;

        #endregion

        #region Property

        public string PinName
        {
            get { return _pinName; }
            set { _pinName = value; }
        }

        public string vmain
        {
            get { return _vmain; }
            set { _vmain = value; }
        }

        public string valt
        {
            get { return _valt; }
            set { _valt = value; }
        }

        public string iFoldLevel
        {
            get { return _iFoldLevel; }
            set { _iFoldLevel = value; }
        }

        public string tdelay
        {
            get { return _tdelay; }
            set { _tdelay = value; }
        }
        #endregion

        #region Construtor

        public PowerLevel(string pinName, string vmain, string valt, string ifoldlevel, string tdelay)
        {
            _pinName = pinName;
            _vmain = vmain;
            _valt = valt;
            _iFoldLevel = ifoldlevel;
            _tdelay = tdelay;
        }

        #endregion

        #region Member Function



        #endregion
    }

    public class DcviPowerLevel
    {
        #region Field
        private string _pinName;
        private string _vps;
        private string _isc;
        private string _tdelay;
        #endregion

        #region Property
        public string PinName
        {
            get { return _pinName; }
            set { _pinName = value; }
        }
        public string vps
        {
            get { return _vps; }
            set { _vps = value; }
        }

        public string isc
        {
            get { return _isc; }
            set { _isc = value; }
        }
        public string tdelay
        {
            get { return _tdelay; }
            set { _tdelay = value; }
        }
        #endregion

        #region Construtor
        public DcviPowerLevel(string pinName, string vps, string isc, string tdelay)
        {
            _pinName = pinName;
            _vps = vps;
            _isc = isc;
            _tdelay = tdelay;
        }
        #endregion

        #region Member Function
        
        #endregion
    }
}