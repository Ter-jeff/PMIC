using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadGlobalSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region Private Function

        private GlobalSpec GetGlobalSpecRow(ExcelWorksheet sheet, int row)
        {
            var symbol = "";
            var job = "";
            var value = "";
            var comment = "";
            for (var i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                var lStrHead = GetMergeCellValue(sheet, StartRowIndex, i);
                var content = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "SYMBOL":
                        symbol = content;
                        break;
                    case "JOB":
                        job = content;
                        break;
                    case "VALUE":
                        value = content;
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }

            return new GlobalSpec(symbol, value, job, comment);
        }

        #endregion

        #region public Function

        public GlobalSpecSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public GlobalSpecSheet GetSheet(ExcelWorksheet sheet)
        {
            var globalSpecSheet = new GlobalSpecSheet(sheet, false);
            var maxRowCount = sheet.Dimension.End.Row;

            // Set Row
            for (var i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                var lDataRow = GetGlobalSpecRow(sheet, i);
                if (string.IsNullOrEmpty(lDataRow.Symbol))
                    break;
                globalSpecSheet.AddRow(lDataRow);
            }

            return globalSpecSheet;
        }

        #endregion
    }
}