using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
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
        public Dictionary<int, string> ItemsWithIndex { get; set; }
        #endregion

        public bool IsEmptyRow
        {
            get
            {
                return string.IsNullOrEmpty(Name) &&
                       string.IsNullOrEmpty(ItemList) &&
                       string.IsNullOrEmpty(Op) &&
                       string.IsNullOrEmpty(Sort) &&
                       string.IsNullOrEmpty(Bin);
            }
        }

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
            ItemsWithIndex = new Dictionary<int, string>();
        }

        public BinTableRow CopyBinTableRow()
        {
            var newitem = new BinTableRow();
            newitem.LinNum = LinNum;
            newitem.Name = Name;
            newitem.ItemList = ItemList;
            newitem.Op = Op;
            newitem.Sort = Sort;
            newitem.Bin = Bin;
            newitem.Result = Result;
            newitem.Comment = Comment;
            newitem.ExtraBinDictionary = new Dictionary<string, string>();
            newitem.Items = new List<string>();
            foreach (var dicItem in ExtraBinDictionary)
            {
                newitem.ExtraBinDictionary.Add(dicItem.Key, dicItem.Value);
            }

            newitem.Items.AddRange(Items);

            //newitem.BinTableEntries.AddRange(BinTableEntries);

            foreach (var dicItem in ItemsWithIndex)
            {
                newitem.ItemsWithIndex.Add(dicItem.Key, dicItem.Value);
            }

            return newitem;
        }
        #endregion
    }
}
