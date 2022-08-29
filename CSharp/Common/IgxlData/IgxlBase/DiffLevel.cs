namespace IgxlData.IgxlBase
{
    public class DiffLevel
    {
        #region Constructor

        public DiffLevel(string pinName, string vicm, string vid, string dVid0, string dVid1, string dVicm0,
            string dVicm1, string vod,
            string vodAlt1, string vodAlt2, string dVod0, string dVod1, string iol, string ioh, string vodTyp,
            string vocmTyp, string vt, string vcl, string vch, string driverMode)
        {
            PinName = pinName;
            Vicm = vicm;
            Vid = vid;
            DVid0 = dVid0;
            DVid1 = dVid1;
            DVicm0 = dVicm0;
            DVicm1 = dVicm1;
            Vod = vod;
            VodAlt1 = vodAlt1;
            VodAlt2 = vodAlt2;
            DVod0 = dVod0;
            DVod1 = dVod1;
            Iol = iol;
            Ioh = ioh;
            VodTyp = vodTyp;
            VocmTyp = vocmTyp;
            Vt = vt;
            Vcl = vcl;
            Vch = vch;
            DriverMode = driverMode;
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string PinName { get; set; }

        public string Vicm { get; set; }

        public string Vid { get; set; }

        public string DVid0 { get; set; }

        public string DVid1 { get; set; }

        public string DVicm0 { get; set; }

        public string DVicm1 { get; set; }

        public string Vod { get; set; }

        public string VodAlt1 { get; set; }

        public string VodAlt2 { get; set; }

        public string DVod0 { get; set; }

        public string DVod1 { get; set; }

        public string Iol { get; set; }

        public string Ioh { get; set; }

        public string VodTyp { get; set; }

        public string VocmTyp { get; set; }

        public string Vt { get; set; }

        public string Vcl { get; set; }

        public string Vch { get; set; }

        public string DriverMode { get; set; }

        #endregion
    }
}