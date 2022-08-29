using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PSet : IgxlItem
    {
        public string Comment;
        public string InstrumentType;
        public string Name;
        public Dictionary<string, string> Parameters;
        public string Pin;
        public string ThislargeheadingmakesAutoFitenlargetherowheight;

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