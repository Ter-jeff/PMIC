namespace IgxlData.IgxlBase
{
    public class PowerLevel
    {
        #region Construtor

        public PowerLevel(string pinName, string vmain, string valt, string ifoldLevel, string tdelay)
        {
            PinName = pinName;
            Vmain = vmain;
            Valt = valt;
            FoldLevel = ifoldLevel;
            Tdelay = tdelay;
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string PinName { get; set; }

        public string Vmain { get; set; }

        public string Valt { get; set; }

        public string FoldLevel { get; set; }

        public string Tdelay { get; set; }

        #endregion

        #region Member Function

        #endregion
    }

    public class DcviPowerLevel
    {
        #region Construtor

        public DcviPowerLevel(string pinName, string vps, string isc, string tdelay)
        {
            PinName = pinName;
            Vps = vps;
            Isc = isc;
            Tdelay = tdelay;
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string PinName { get; set; }

        public string Vps { get; set; }

        public string Isc { get; set; }

        public string Tdelay { get; set; }

        #endregion

        #region Member Function

        #endregion
    }

    public class Dc30Level
    {
        #region Construtor

        public Dc30Level(string pinName, string vlevel)
        {
            PinName = pinName;
            Vlevel = vlevel;
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string PinName { get; set; }

        public string Vlevel { get; set; }

        #endregion

        #region Member Function

        #endregion
    }
}