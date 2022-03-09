using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadAcSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 7;

        #region public Function
        public AcSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public AcSpecSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public AcSpecSheet GetSheet(ExcelWorksheet sheet)
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
            var acSpecSheet = new AcSpecSheet(sheet, categoryList, selectorNameList);
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var symbol = GetMergeCellValue(sheet, i, 2).Trim();
                if (string.IsNullOrEmpty(symbol)) break;
                var lDataRow = GetAcSpecsRow(sheet, i);
                acSpecSheet.AddRow(lDataRow);
            }
            return acSpecSheet;
        }
        #endregion

        #region Private Function
        private AcSpecs GetAcSpecsRow(ExcelWorksheet sheet, int row)
        {
            var symbol = GetMergeCellValue(sheet, row, 2).Trim();
            var name = "";
            var comment = "";
            var typ = "";
            var min = "";
            var max = "";
            var categroyItems = new List<CategoryInSpec>();
            var selectorList = new List<Selector>();
            for (var i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                var lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead))
                {
                    if (!string.IsNullOrEmpty(name))
                        categroyItems.Add(new CategoryInSpec(name, typ, min, max));
                    name = lStrHead;
                }
                var lStrHead2 = GetMergeCellValue(sheet, StartRowIndex + 1, i).Trim();
                var lStrContent = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead2))
                {
                    case "TYP":
                        typ = lStrContent;
                        selectorList.Add(new Selector("Typ", "Typ"));
                        break;
                    case "MIN":
                        min = lStrContent;
                        selectorList.Add(new Selector("Min", "Min"));
                        break;
                    case "MAX":
                        max = lStrContent;
                        selectorList.Add(new Selector("Max", "Max"));
                        break;
                    case "COMMENT":
                        comment = lStrContent;
                        break;
                }
            }

            categroyItems.Add(new CategoryInSpec(name, typ, min, max));
            var acspecs = new AcSpecs(symbol, selectorList, "", comment);
            foreach (var categroyItem in categroyItems)
                acspecs.AddCategory(categroyItem);
            return acspecs;
        }
        #endregion
    }
}
