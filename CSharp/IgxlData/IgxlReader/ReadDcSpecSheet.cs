using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadDcSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 7;

        #region public Function
        public DcSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public DcSpecSheet GetSheet(ExcelWorksheet sheet)
        {
            List<string> categoryList = new List<string>();
            List<string> selectorNameList = new List<string>();
            int maxRowCount = sheet.Dimension.End.Row;
            int maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                string lStrHead2 = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead))
                    categoryList.Add(lStrHead);
                if (!string.IsNullOrEmpty(lStrHead2))
                    selectorNameList.Add(lStrHead2);
            }

            // Set Row
            DcSpecSheet dcSpecSheet = new DcSpecSheet(sheet, categoryList, selectorNameList);
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                string symbol = GetMergeCellValue(sheet, i, 2).Trim();
                if (string.IsNullOrEmpty(symbol)) break;
                DcSpecs lDataRow = GetDcSpecsRow(sheet, i);
                dcSpecSheet.AddRow(lDataRow);
            }
            return dcSpecSheet;
        }
        #endregion

        #region Private Function
        private DcSpecs GetDcSpecsRow(ExcelWorksheet sheet, int row)
        {
            string symbol = GetMergeCellValue(sheet, row, 2).Trim();
            string name = "";
            string comment = "";
            string typ = "";
            string min = "";
            string max = "";
            List<CategoryInSpec> categoryInSpecs = new List<CategoryInSpec>();
            List<Selector> selectorList = new List<Selector>();
            for (int i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                string lStrHead = GetMergeCellValue(sheet, StartRowIndex, i).Trim();
                if (!string.IsNullOrEmpty(lStrHead))
                {
                    if (!string.IsNullOrEmpty(name))
                        categoryInSpecs.Add(new CategoryInSpec(name, typ, min, max));
                    name = lStrHead;
                }
                string lStrHead2 = GetMergeCellValue(sheet, StartRowIndex + 1, i).Trim();
                string content = GetCellText(sheet, row, i);
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
            var dcSpecs = new DcSpecs(symbol, selectorList, "", comment);
            foreach (var categoryInSpec in categoryInSpecs)
                dcSpecs.AddCategory(categoryInSpec);
            return dcSpecs;
        }
        #endregion
    }
}
