using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Library.DataStruct
{
    public class ItemInfo
    {
        public string Name;
        public List<Regex> Patterns = new List<Regex>();
    }
}
