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
        #region Property

        public int LinNum { get; set; }
        public string Name { get; set; }
        public string ItemList { get; set; }
        public string Op { get; set; }
        public string Sort { get; set; }
        public string Bin { get; set; }
        public Dictionary<string, string> ExtraBinDictionary { get; set; }
        public string Result { get; set; }
        public string Comment { get; set; }
        public List<string> Items { get; set; }

        #endregion

        #region Constructor

        public BinTableRow()
        {
            LinNum = 0;
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

        public BinTableRow DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as BinTableRow;
            }
        }

        public BinTableRow CopyBinTableRow()
        {
            var binTableRow = new BinTableRow();
            binTableRow.LinNum = LinNum;
            binTableRow.Name = Name;
            binTableRow.ItemList = ItemList;
            binTableRow.Op = Op;
            binTableRow.Sort = Sort;
            binTableRow.Bin = Bin;
            binTableRow.Result = Result;
            binTableRow.Comment = Comment;
            binTableRow.ExtraBinDictionary = new Dictionary<string, string>();
            binTableRow.Items = new List<string>();
            foreach (var dicItem in ExtraBinDictionary) binTableRow.ExtraBinDictionary.Add(dicItem.Key, dicItem.Value);

            foreach (var listItem in Items) binTableRow.Items.Add(listItem);
            return binTableRow;
        }

        #endregion
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