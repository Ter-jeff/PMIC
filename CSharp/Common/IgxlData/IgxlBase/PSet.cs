using System.Collections.Generic;
using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]
    public class PSet : IgxlRow
    {
        public string InstrumentType { get; set; }
        public string Name { get; set; }

        public Dictionary<string, string> Parameters;
        public string Pin { get; set; }
        public string ThislargeheadingmakesAutoFitenlargetherowheight { get; set; }
        public string Comment { get; set; }

        public PSet()
        {
            Name = "";
            Pin = "";
            InstrumentType = "";
            Parameters = new Dictionary<string, string>();
            ThislargeheadingmakesAutoFitenlargetherowheight = "";
            Comment = "";
        }
    }
}