using System.Collections.Generic;
using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]
    public class PSet : IgxlRow
    {
        public string InstrumentType;
        public string Name;
        public Dictionary<string, string> Parameters;
        public string Pin;
        public string ThislargeheadingmakesAutoFitenlargetherowheight;
        public string Comment;

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