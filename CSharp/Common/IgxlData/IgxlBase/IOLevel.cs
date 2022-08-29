namespace IgxlData.IgxlBase
{
    public class IoLevel
    {
        #region Constructor

        public IoLevel(string pinName, string vil, string vih, string vol, string voh, string vohAlt1, string vohAtl2,
            string iol,
            string ioh, string vt, string vcl, string vch, string voutLoTyp, string voutHiTyp, string driverMode)
        {
            PinName = pinName;
            Vil = vil;
            Vih = vih;
            Vol = vol;
            Voh = voh;
            VohAlt1 = vohAlt1;
            VohAlt2 = vohAtl2;
            Iol = iol;
            Ioh = ioh;
            Vt = vt;
            Vcl = vcl;
            Vch = vch;
            VoutLoTyp = voutLoTyp;
            VoutHiTyp = voutHiTyp;
            DriverMode = driverMode;
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string PinName { get; set; }

        public string Vil { get; set; }

        public string Vih { get; set; }

        public string Vol { get; set; }

        public string Voh { get; set; }

        public string VohAlt1 { get; set; }

        public string VohAlt2 { get; set; }

        public string Iol { get; set; }

        public string Ioh { get; set; }

        public string Vt { get; set; }

        public string Vcl { get; set; }

        public string Vch { get; set; }

        public string VoutLoTyp { get; set; }

        public string VoutHiTyp { get; set; }

        public string DriverMode { get; set; }

        #endregion
    }
}