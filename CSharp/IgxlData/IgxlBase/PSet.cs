using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PSet:IgxlItem
    {
        public string Name;
        public string Pin;
        public string InstrumentType;
        public Dictionary<string, string> Parameters;
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
