
namespace IgxlData.IgxlBase
{
    public class DiffLevel
    {
        #region Field

        private string _pinName;
        private string _Vicm;
        private string _Vid;
        private string _dVid0;
        private string _dVid1;
        private string _dVicm0;
        private string _dVicm1;
        private string _Vod;
        private string _Vod_Alt1;
        private string _Vod_Alt2;
        private string _dVod0;
        private string _dVod1;
        private string _Iol;
        private string _Ioh;
        private string _VodTyp;
        private string _VocmTyp;
        private string _Vt;
        private string _Vcl;
        private string _Vch;
        private string _DriverMode;

        #endregion

        #region Property
        public string PinName
        {
            get { return _pinName; }
            set { _pinName = value; }
        }
        public string Vicm
        {
            get { return _Vicm; }
            set { _Vicm = value; }
        }
        public string Vid
        {
            get { return _Vid; }
            set { _Vid = value; }
        }
        public string DVid0
        {
            get { return _dVid0; }
            set { _dVid0 = value; }
        }
        public string DVid1
        {
            get { return _dVid1; }
            set { _dVid1 = value; }
        }
        public string DVicm0
        {
            get { return _dVicm0; }
            set { _dVicm0 = value; }
        }
        public string DVicm1
        {
            get { return _dVicm1; }
            set { _dVicm1 = value; }
        }
        public string Vod
        {
            get { return _Vod; }
            set { _Vod = value; }
        }
        public string Vod_Alt1
        {
            get { return _Vod_Alt1; }
            set { _Vod_Alt1 = value; }
        }
        public string Vod_Alt2
        {
            get { return _Vod_Alt2; }
            set { _Vod_Alt2 = value; }
        }
        public string DVod0
        {
            get { return _dVod0; }
            set { _dVod0 = value; }
        }
        public string DVod1
        {
            get { return _dVod1; }
            set { _dVod1 = value; }
        }
        public string Iol
        {
            get { return _Iol; }
            set { _Iol = value; }
        }
        public string Ioh
        {
            get { return _Ioh; }
            set { _Ioh = value; }
        }
        public string VodTyp
        {
            get { return _VodTyp; }
            set { _VodTyp = value; }
        }
        public string VocmTyp
        {
            get { return _VocmTyp; }
            set { _VocmTyp = value; }
        }
        public string Vt
        {
            get { return _Vt; }
            set { _Vt = value; }
        }
        public string Vcl
        {
            get { return _Vcl; }
            set { _Vcl = value; }
        }
        public string Vch
        {
            get { return _Vch; }
            set { _Vch = value; }
        }
        public string DriverMode
        {
            get { return _DriverMode; }
            set { _DriverMode = value; }
        }
        #endregion

        #region Constructor

        public DiffLevel(string PinName, string Vicm, string Vid, string dVid0, string dVid1, string dVicm0, string dVicm1, string Vod,
                            string Vod_Alt1, string Vod_Alt2, string dVod0, string dVod1, string Iol, string Ioh, string VodTyp,
                            string VocmTyp, string Vt, string Vcl, string Vch, string DriverMode)
        {
            _pinName = PinName;
            _Vicm = Vicm;
            _Vid = Vid;
            _dVid0 = dVid0;
            _dVid1 = dVid1;
            _dVicm0 = dVicm0;
            _dVicm1 = dVicm1;
            _Vod = Vod;
            _Vod_Alt1 = Vod_Alt1;
            _Vod_Alt2 = Vod_Alt2;
            _dVod0 = dVod0;
            _dVod1 = dVod1;
            _Iol = Iol;
            _Ioh = Ioh;
            _VodTyp = VodTyp;
            _VocmTyp = VocmTyp;
            _Vt = Vt;
            _Vcl = Vcl;
            _Vch = Vch;
            _DriverMode = DriverMode;
        }

        #endregion
    }
}