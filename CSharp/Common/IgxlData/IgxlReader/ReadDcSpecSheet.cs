using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadDcSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 7;

        #region Private Function

        private DcSpec GetDcSpecsRow(ExcelWorksheet sheet, int row)
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

        #endregion

        #region public Function

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
                var lDataRow = GetDcSpecsRow(sheet, i);
                dcSpecSheet.AddRow(lDataRow);
            }

            return dcSpecSheet;
        }

        #endregion
    }
}