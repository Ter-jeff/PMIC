using System;
using System.Collections.Generic;
using IgxlData.IgxlBase;

namespace IgxlData.Others.MultiTimeSet
{
    [Serializable]
    public class ComTimeSetBasic : Tset
    {
        public Dictionary<string, double> ShiftInReserve;

        public Dictionary<string, double>
            SubCommentVariable; // = new Dictionary<string, double>();  //Store var from comment under TSB sheet

        public List<string>
            SubContextVariable; // = new List<string>();  //Store var from context in TSB sheet, save 2 parts separately in order to judge un-match variable

        public ComTimeSetBasic()
        {
            SubCommentVariable = new Dictionary<string, double>();
            SubContextVariable = new List<string>();
            ShiftInReserve = new Dictionary<string, double>();
        }
    }
}