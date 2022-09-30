using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class PowerLevel
    {
        public PowerLevel(string pinName, string vmain, string valt, string ifoldLevel, string tdelay)
        {
            PinName = pinName;
            Vmain = vmain;
            Valt = valt;
            FoldLevel = ifoldLevel;
            Tdelay = tdelay;
        }

        public string PinName { get; set; }
        public string Vmain { get; set; }
        public string Valt { get; set; }
        public string FoldLevel { get; set; }
        public string Tdelay { get; set; }

    }
}