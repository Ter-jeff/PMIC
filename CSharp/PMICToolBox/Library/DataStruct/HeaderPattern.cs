using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class HeaderPattern
    {
        public string Name;
        public string Pattern;
        public List<HeaderItem> Items = new List<HeaderItem>();
        public Regex HeaderReg;
        public Regex DataRegex;
    }

    public class HeaderItem
    {
        public string Name;
        public bool Missingpossible;
        public string Pattern;
    }
}
