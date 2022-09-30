using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class Dc30Level
    {
        public Dc30Level(string pinName, string vlevel)
        {
            PinName = pinName;
            Vlevel = vlevel;
        }

        public string PinName { get; set; }
        public string Vlevel { get; set; }

    }
}