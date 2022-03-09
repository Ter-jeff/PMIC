using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PSet : IgxlRow
    {
        public string Name;
        public string Pin;
        public string InstrumentType;
        public string ThislargeheadingmakesAutoFitenlargetherowheight;
        public Dictionary<string, string> Parameters;
        public string Comment;


        public PSet()
        {
            Name = "";
            Pin = "";
            InstrumentType = "";
            ThislargeheadingmakesAutoFitenlargetherowheight = "";
            Parameters = new Dictionary<string, string>();
            Comment = "";
        }
    }
}
