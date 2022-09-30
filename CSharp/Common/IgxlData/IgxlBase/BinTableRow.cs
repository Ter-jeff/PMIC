using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace IgxlData.IgxlBase
{
    [Serializable]
    [DebuggerDisplay("{Name}")]
    public class BinTableRow : IgxlRow
    {
        public string Name { get; set; }
        public string ItemList { get; set; }
        public string Op { get; set; }
        public string Sort { get; set; }
        public string Bin { get; set; }
        public Dictionary<string, string> ExtraBinDictionary { get; set; }
        public string Result { get; set; }
        public string Comment { get; set; }
        public List<string> Items { get; set; }

        public BinTableRow()
        {
            Name = "";
            ItemList = "";
            Op = "";
            Sort = "";
            Bin = "";
            Result = "";
            Comment = "";
            ExtraBinDictionary = new Dictionary<string, string>();
            Items = new List<string>();
        }
    }

    public class BinTableRowComparer : IEqualityComparer<BinTableRow>
    {
        public bool Equals(BinTableRow x, BinTableRow y)
        {
            if (x == null || y == null)
                return false;
            if (x.Name == null || y.Name == null)
                return false;

            if (x.Name.Equals(y.Name, StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        public int GetHashCode(BinTableRow obj)
        {
            if (obj == null)
                return 0;
            if (obj.Name == null)
                return 0;

            return obj.Name.ToLower().GetHashCode();
        }
    }
}