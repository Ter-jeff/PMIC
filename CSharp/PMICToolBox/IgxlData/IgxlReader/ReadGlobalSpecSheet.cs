using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadGlobalSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public GlobalSpecSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public GlobalSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public GlobalSpecSheet GetSheet(ExcelWorksheet sheet)
        {
            GlobalSpecSheet globalSpecSheet = new GlobalSpecSheet(sheet, false);
            int maxRowCount = sheet.Dimension.End.Row;
            int maxColumnCount = sheet.Dimension.End.Column;

            // Set Row
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                GlobalSpec lDataRow = GetGlobalSpecRow(sheet, i);
                if (string.IsNullOrEmpty(lDataRow.Symbol))
                    break;
                globalSpecSheet.AddRow(lDataRow);
            }
            return globalSpecSheet;
        }
        #endregion

        #region Private Function
        private GlobalSpec GetGlobalSpecRow(ExcelWorksheet sheet, int row)
        {
            string symbol = "";
            string job = "";
            string value = "";
            string comment = "";
            for (int i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                string lStrHead = GetMergeCellValue(sheet, StartRowIndex, i);
                string lStrContent = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "SYMBOL":
                        symbol = lStrContent;
                        break;
                    case "JOB":
                        job = lStrContent;
                        break;
                    case "VALUE":
                        value = lStrContent;
                        break;
                    case "COMMENT":
                        comment = lStrContent;
                        break;
                }
            }
            return new GlobalSpec(symbol, value, job, comment);
        }
        #endregion
    }
}
