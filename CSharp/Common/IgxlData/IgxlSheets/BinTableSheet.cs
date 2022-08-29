using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using OfficeOpenXml;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class BinTableSheet : IgxlSheet
    {
        private const string SheetType = "DTBintablesSheet";
        public BinTableRows BinTableRows { get; set; }

        #region Constructor

        public BinTableSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            BinTableRows = new BinTableRows();
            IgxlSheetName = IgxlSheetNameList.BinTable;
        }

        public BinTableSheet(string sheetName)
            : base(sheetName)
        {
            BinTableRows = new BinTableRows();
            IgxlSheetName = IgxlSheetNameList.BinTable;
        }

        #endregion

        #region Member Function

        public void AddRow(BinTableRow row)
        {
            BinTableRows.Add(row);
        }

        public void AddRows(BinTableRows rows)
        {
            BinTableRows.AddRange(rows);
        }

        public void RemoveRow(BinTableRow row)
        {
            BinTableRows.Remove(row);
        }

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
        }


        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.0";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet2_0(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The binTable sheet version:{0} is not supported!", version));
                }
            }
        }

        private void WriteSheet2_0(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (BinTableRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                const string pattern = @"^Item(?<number>[\d]+)";
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var nameIndex = GetIndexFrom(igxlSheetsVersion, "Name");
                var itemListIndex = GetIndexFrom(igxlSheetsVersion, "Item List");
                var opIndex = GetIndexFrom(igxlSheetsVersion, "Op");
                var sortIndex = GetIndexFrom(igxlSheetsVersion, "Sort");
                var binIndex = GetIndexFrom(igxlSheetsVersion, "Bin");
                var resultIndex = GetIndexFrom(igxlSheetsVersion, "Result");
                var item0Index = GetIndexFrom(igxlSheetsVersion, "Item0");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
                             IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    foreach (var item in igxlSheetsVersion.Columns.Column)
                        if (item.indexFrom == item.indexTo && item.rowIndex == i)
                            if (Regex.IsMatch(item.columnName, pattern))
                                arr[item.indexFrom] = item.columnName.Replace("Item", "Items");
                            else
                                arr[item.indexFrom] = item.columnName;

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < BinTableRows.Count; index++)
                {
                    var row = BinTableRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Name))
                    {
                        arr[0] = row.ColumnA;
                        arr[nameIndex] = row.Name;
                        arr[itemListIndex] = string.Join(",", row.ItemList);
                        arr[opIndex] = row.Op;
                        arr[sortIndex] = row.Sort;
                        arr[binIndex] = row.Bin;
                        arr[resultIndex] = row.Result;
                        for (var i = 0; i < row.Items.Count; i++)
                        {
                            var item = row.Items[i];
                            arr[item0Index + i] = item;
                        }

                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] {"\t"};
                    }

                    sw.WriteLine(string.Join("\t", arr));
                }

                #endregion
            }
        }

        #endregion
    }
}