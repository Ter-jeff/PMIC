using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadDcSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int EndRowIndex = 4;
        private const int StartColumnIndex = 7;

        public DcSpecSheet GetSheet(Stream stream, string sheetName)
        {
            var dcSpecSheet = new DcSpecSheet(sheetName);
            var isBackup = false;
            var i = 1;
            var categoryLine = "";
            var categoryList = new List<string>();
            var selectorLine = "";
            var selectorNameList = new List<string>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > EndRowIndex)
                    {
                        var acSpec = GetDcSpecsRow(line, sheetName, i, selectorNameList,
                            categoryLine, selectorLine);
                        if (string.IsNullOrEmpty(acSpec.Symbol))
                        {
                            isBackup = true;
                            continue;
                        }

                        acSpec.IsBackup = isBackup;
                        dcSpecSheet.AddRow(acSpec);
                    }
                    else
                    {
                        var arr = line.Split('\t');
                        var maxColumnCount = arr.Length;

                        if (i == StartRowIndex)
                        {
                            for (var col = StartColumnIndex; col < maxColumnCount; col++)
                            {
                                var value = arr[col];
                                if (!string.IsNullOrEmpty(value) && col != StartColumnIndex)
                                {
                                    categoryList.Add(value);
                                    categoryLine = line;
                                }
                            }
                        }
                        else if (i == StartRowIndex + 1)
                        {
                            for (var col = StartColumnIndex; col < maxColumnCount; col++)
                            {
                                var value = arr[col];
                                if (!string.IsNullOrEmpty(value))
                                {
                                    selectorNameList.Add(value);
                                    selectorLine = line;
                                }
                            }

                            selectorNameList = selectorNameList.Distinct().ToList();
                            selectorNameList.Remove("Comment");
                        }
                    }
                    i++;
                }
            }
            dcSpecSheet.CategoryList = categoryList;
            dcSpecSheet.SelectorNameList = selectorNameList;
            return dcSpecSheet;
        }

        private DcSpec GetDcSpecsRow(string line, string sheetName, int row, List<string> selectorNameList,
            string categoryLine, string selectorLine)
        {
            var arr = line.Split('\t');
            var arrCategory = categoryLine.Split('\t');
            var arrSelector = selectorLine.Split('\t');

            var symbol = arr[1];
            var name = "";
            var comment = "";
            var typ = "";
            var min = "";
            var max = "";
            var categoryInSpecs = new List<CategoryInSpec>();
            var selectorList = new List<Selector>();
            for (var i = StartColumnIndex + selectorNameList.Count - 1; i < arr.Length; i++)
            {
                var category = arrCategory[i].Trim();
                if (!string.IsNullOrEmpty(category))
                {
                    if (!string.IsNullOrEmpty(name))
                        categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
                    name = category;
                }

                var selectorName = arrSelector[i].Trim();
                var value = arr[i].Trim();
                switch (FormatStringForCompare(selectorName))
                {
                    case "TYP":
                        typ = value;
                        selectorList.Add(new Selector("Typ", "Typ"));
                        break;
                    case "MIN":
                        min = value;
                        selectorList.Add(new Selector("Min", "Min"));
                        break;
                    case "MAX":
                        max = value;
                        selectorList.Add(new Selector("Max", "Max"));
                        break;
                    case "COMMENT":
                        comment = value;
                        break;
                }
            }

            categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
            var dcSpec = new DcSpec(symbol, selectorList, "", comment);
            dcSpec.RowNum = row;
            dcSpec.SheetName = sheetName;
            foreach (var categoryInSpec in categoryInSpecs)
                dcSpec.AddCategory(categoryInSpec);
            return dcSpec;
        }

        private DcSpec GetDcSpecRow(ExcelWorksheet sheet, int row)
        {
            var symbol = GetMergeCellValue(sheet, row, 2).Trim();
            var name = "";
            var comment = "";
            var typ = "";
            var min = "";
            var max = "";
            var categoryInSpecs = new List<CategoryInSpec>();
            var selectorList = new List<Selector>();
            for (var i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                var lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead))
                {
                    if (!string.IsNullOrEmpty(name))
                        categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
                    name = lStrHead;
                }

                var lStrHead2 = GetMergeCellValue(sheet, StartRowIndex + 1, i).Trim();
                var content = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead2))
                {
                    case "TYP":
                        typ = content;
                        selectorList.Add(new Selector("Typ", "Typ"));
                        break;
                    case "MIN":
                        min = content;
                        selectorList.Add(new Selector("Min", "Min"));
                        break;
                    case "MAX":
                        max = content;
                        selectorList.Add(new Selector("Max", "Max"));
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }

            categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
            var dcSpecs = new DcSpec(symbol, selectorList, "", comment);
            foreach (var categoryInSpec in categoryInSpecs)
                dcSpecs.AddCategory(categoryInSpec);
            return dcSpecs;
        }

        public DcSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public DcSpecSheet GetSheet(ExcelWorksheet sheet)
        {
            var categoryList = new List<string>();
            var selectorNameList = new List<string>();
            var maxRowCount = sheet.Dimension.End.Row;
            var maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (var i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                var lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                var lStrHead2 = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead))
                    categoryList.Add(lStrHead);
                if (!string.IsNullOrEmpty(lStrHead2))
                    selectorNameList.Add(lStrHead2);
            }

            // Set Row
            var dcSpecSheet = new DcSpecSheet(sheet, categoryList, selectorNameList);
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var symbol = GetMergeCellValue(sheet, i, 2).Trim();
                if (string.IsNullOrEmpty(symbol)) break;
                var lDataRow = GetDcSpecRow(sheet, i);
                dcSpecSheet.AddRow(lDataRow);
            }

            return dcSpecSheet;
        }
    }
}