using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
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

        public AcSpecSheet GetSheet(ExcelWorksheet sheet)
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
            AcSpecSheet acSpecSheet = new AcSpecSheet(sheet, categoryList, selectorNameList);
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                string symbol = GetMergeCellValue(sheet, i, 2).Trim();
                if (string.IsNullOrEmpty(symbol)) break;
                AcSpecs lDataRow = GetAcSpecsRow(sheet, i);
                acSpecSheet.AddRow(lDataRow);
            }
            return acSpecSheet;
        }
        #endregion

        #region Private Function
        private AcSpecs GetAcSpecsRow(ExcelWorksheet sheet, int row)
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
            var acSpecs = new AcSpecs(symbol, selectorList, "", comment);
            foreach (var categoryInSpec in categoryInSpecs)
                acSpecs.AddCategory(categoryInSpec);
            return acSpecs;
        }
        #endregion
    }
}
