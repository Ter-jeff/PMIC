using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadAcSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int EndRowIndex = 4;
        private const int StartColumnIndex = 4;
        private const int StartColumn1Index = 6;
        private const int StartColumn2Index = 3;

        public AcSpecSheet GetSheet(Stream stream, string sheetName)
        {
            var acSpecSheet = new AcSpecSheet(sheetName);
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
                        var acSpec = GetAcSpecsRow(line, sheetName, i, selectorNameList,
                            categoryLine, selectorLine);
                        if (string.IsNullOrEmpty(acSpec.Symbol))
                        {
                            isBackup = true;
                            continue;
                        }

                        acSpec.IsBackup = isBackup;
                        acSpecSheet.AddRow(acSpec);
                        if (acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                            foreach (var category in acSpec.CategoryList)
                                if (!string.IsNullOrEmpty(category.Typ))
                                {
                                    var timeSet = category.Typ;
                                    if (!acSpecSheet.CategoryTimeSetDic.ContainsKey(category.Name))
                                        acSpecSheet.CategoryTimeSetDic.Add(category.Name, new List<string> { timeSet });
                                    else if (!acSpecSheet.CategoryTimeSetDic[category.Name].ToList()
                                                 .Contains(timeSet, StringComparer.CurrentCultureIgnoreCase))
                                        acSpecSheet.CategoryTimeSetDic[category.Name].Add(timeSet);
                                }
                    }
                    else
                    {
                        var arr = line.Split('\t');
                        var maxColumnCount = arr.Length;

                        if (i == StartRowIndex)
                        {
                            for (var col = StartColumn1Index; col < maxColumnCount; col++)
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
                            for (var col = StartColumn2Index; col < maxColumnCount; col++)
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
            acSpecSheet.CategoryList = categoryList;
            acSpecSheet.SelectorNameList = selectorNameList;
            return acSpecSheet;
        }

        private AcSpec GetAcSpecsRow(string line, string sheetName, int row, List<string> selectorNameList,
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
            var acSpec = new AcSpec(symbol, selectorList, "", comment);
            acSpec.RowNum = row;
            acSpec.SheetName = sheetName;
            foreach (var categoryInSpec in categoryInSpecs)
                acSpec.AddCategory(categoryInSpec);
            return acSpec;
        }

        private AcSpec GetAcSpecsRow(ExcelWorksheet sheet, int row, List<string> selectorNameList)
        {
            var symbol = GetMergeCellValue(sheet, row, 2).Trim();
            var name = "";
            var comment = "";
            var typ = "";
            var min = "";
            var max = "";
            var categoryInSpecs = new List<CategoryInSpec>();
            var selectorList = new List<Selector>();
            for (var i = StartColumnIndex + selectorNameList.Count; i < sheet.Dimension.End.Column; i++)
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
            var acSpec = new AcSpec(symbol, selectorList, "", comment);
            acSpec.RowNum = row;
            acSpec.SheetName = sheet.Name;
            foreach (var categoryInSpec in categoryInSpecs)
                acSpec.AddCategory(categoryInSpec);
            return acSpec;
        }

        public AcSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public AcSpecSheet GetSheet(ExcelWorksheet sheet)
        {
            var categoryList = new List<string>();
            var selectorNameList = new List<string>();
            var maxRowCount = sheet.Dimension.End.Row;
            var maxColumnCount = sheet.Dimension.End.Column;

            var stop = false;
            for (var i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                var lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                var lStrHead2 = GetMergeCellValue(sheet, StartRowIndex + 1, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead) && i != StartColumnIndex)
                {
                    categoryList.Add(lStrHead);
                    stop = true;
                }

                if (!string.IsNullOrEmpty(lStrHead2) && stop == false)
                    selectorNameList.Add(lStrHead2);
            }

            var acSpecSheet = new AcSpecSheet(sheet, categoryList, selectorNameList);
            var isBackup = false;
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var symbol = GetMergeCellValue(sheet, i, 2).Trim();
                if (string.IsNullOrEmpty(symbol))
                {
                    isBackup = true;
                    continue;
                }
                var row = GetAcSpecsRow(sheet, i, selectorNameList);
                row.IsBackup = isBackup;
                acSpecSheet.AddRow(row);
                if (row.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                    foreach (var category in row.CategoryList)
                        if (!string.IsNullOrEmpty(category.Typ))
                        {
                            var timeSet = category.Typ;
                            if (!acSpecSheet.CategoryTimeSetDic.ContainsKey(category.Name))
                                acSpecSheet.CategoryTimeSetDic.Add(category.Name, new List<string> { timeSet });
                            else if (!acSpecSheet.CategoryTimeSetDic[category.Name].ToList()
                                         .Contains(timeSet, StringComparer.CurrentCultureIgnoreCase))
                                acSpecSheet.CategoryTimeSetDic[category.Name].Add(timeSet);
                        }
            }

            return acSpecSheet;
        }
    }
}