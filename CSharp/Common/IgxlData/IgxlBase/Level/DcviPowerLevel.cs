using System.Diagnostics;

namespace IgxlData.IgxlBase
{

    [DebuggerDisplay("{PinName}")]
    public class DcviPowerLevel
    {
        public DcviPowerLevel(string pinName, string vps, string isc, string tdelay)
        {
            PinName = pinName;
            Vps = vps;
            Isc = isc;
            Tdelay = tdelay;
        }

        public string PinName { get; set; }
        public string Vps { get; set; }
        public string Isc { get; set; }
        public string Tdelay { get; set; }
    }
}