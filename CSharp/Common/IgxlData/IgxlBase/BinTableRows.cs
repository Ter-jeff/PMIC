using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class BinTableRows : List<BinTableRow>
    {
        public BinTableRows()
        {
        }

        public BinTableRows(List<BinTableRow> collection) : base(collection)
        {
        }

        public void GenBlockBinTable(string block)
        {
            var binTableRow = new BinTableRow();
            binTableRow.Name = "Bin_" + block;
            binTableRow.ItemList = "F_" + block;
            binTableRow.Op = "AND";
            binTableRow.Sort = "9999";
            binTableRow.Bin = "9";
            binTableRow.Result = "Fail";
            binTableRow.Items.Add("T");
            Add(binTableRow);
        }

        public void GenSetError(string block)
        {
            var binTableRow = new BinTableRow();
            binTableRow.Name = "Bin_SET_ERROR_" + block;
            binTableRow.ItemList = "F_SET_ERROR_" + block;
            binTableRow.Op = "AND";
            binTableRow.Sort = "9999";
            binTableRow.Bin = "9";
            binTableRow.Result = "Fail";
            binTableRow.Items.Add("T");
            Add(binTableRow);
        }
    }
}