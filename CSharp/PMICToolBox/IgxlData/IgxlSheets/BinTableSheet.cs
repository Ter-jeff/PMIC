using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlSheets
{
    public class BinTableSheet : IgxlSheet
    {
        private const string SheetType = "DTBintablesSheet";

        #region Field
        private List<BinTableRow> _binTableRows;
        #endregion

        #region Constructor
        public BinTableSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _binTableRows = new List<BinTableRow>();
            IgxlSheetName = IgxlSheetNameList.BinTable;
        }

        public BinTableSheet(string sheetName)
            : base(sheetName)
        {
            _binTableRows = new List<BinTableRow>();
            IgxlSheetName = IgxlSheetNameList.BinTable;
        }
        #endregion

        #region Member Function
        public void AddRow(BinTableRow row)
        {
            _binTableRows.Add(row);
        }

        public void RemoveRow(BinTableRow row)
        {
            _binTableRows.Remove(row);
        }

        public List<BinTableRow> BinTableRows
        {
            get { return _binTableRows; }
            set { _binTableRows = value; }
        }

        public override void Write(string fileName, string version = "2.0")
        {
            //if (version == "2.0")
            //{
            //    var validate = new Action<string>((a) => { });
            //    var genBinTaleSheet = new GenBinTaleSheet(fileName, validate, true);
            //    foreach (var binTableRow in _binTableRows)
            //    {
            //        var cnt = 80 - binTableRow.Items.Count;
            //        var list = binTableRow.Items;
            //        for (var i = 0; i < cnt; i++)
            //            list.Add("");
            //        list.Add(binTableRow.Comment);
            //        genBinTaleSheet.AddEntry(binTableRow.Name, binTableRow.ItemList, binTableRow.Op, binTableRow.Sort,
            //            binTableRow.Bin, binTableRow.Result, list.ToArray());
            //    }
            //    genBinTaleSheet.WriteSheet();
            //}
            //else
            //{
            //    throw new Exception(string.Format("The bintable sheet version:{0} is not supported!", version));
            //}

            //Support 2.0
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("2.0"))
                {
                    var igxlSheetsVersion = dic["2.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_binTableRows.Count == 0) return;

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
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    foreach (var item in igxlSheetsVersion.Columns.Column)
                    {
                        if (item.indexFrom == item.indexTo && item.rowIndex == i)
                            if (Regex.IsMatch(item.columnName, pattern))
                                arr[item.indexFrom] = item.columnName.Replace("Item", "Items");
                            else
                                arr[item.indexFrom] = item.columnName;
                    }

                    WriteHeader(arr, sw);
                }
                #endregion

                #region data
                for (var index = 0; index < _binTableRows.Count; index++)
                {
                    var row = _binTableRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Name))
                    {
                        arr[0] = row.ColumnA;
                        arr[nameIndex] = row.Name;
                        arr[itemListIndex] = string.Join(",", row.ItemList);
                        arr[opIndex] =  row.Op;
                        arr[sortIndex] = row.Sort;
                        arr[binIndex] = row.Bin;
                        arr[resultIndex] = row.Result;
                        for (int i = 0; i < row.Items.Count; i++)
                        {
                            var item = row.Items[i];
                            arr[item0Index + i] = item;
                        }
                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] { "\t" };
                    }
                    sw.WriteLine(string.Join("\t", arr));
                }
                #endregion
            }
        }
        #endregion
    }
}