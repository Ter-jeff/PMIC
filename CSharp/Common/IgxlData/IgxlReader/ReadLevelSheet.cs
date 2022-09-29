using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadLevelSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;
        private readonly List<string> _headList = new List<string>();

        #region Private Function

        private LevelRow GetLevelRow(ExcelWorksheet sheet, int row)
        {
            var pinName = "";
            var parameter = "";
            var value = "";
            var comment = "";
            for (var i = StartColumnIndex; i < sheet.Dimension.End.Column; i++)
            {
                var lStrHead = _headList[i - StartColumnIndex];
                var content = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "PIN/GROUP":
                        pinName = content;
                        break;
                    case "SEQ.":
                        break;
                    case "PARAMETER":
                        parameter = content;
                        break;
                    case "VALUE":
                        value = content.Replace("=", "");
                        break;
                    case "COMMENT":
                        comment = content;
                        break;
                }
            }

            return new LevelRow(pinName, parameter, value, comment);
        }

        #endregion

        #region public Function

        public LevelSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public LevelSheet GetSheet(ExcelWorksheet sheet)
        {
            var levelSheet = new LevelSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;
            var maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (var i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                var lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                var lStrHead = lStrValue.Trim();
                _headList.Add(lStrHead);
            }

            // Set Row
            for (var i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                var lDataRow = GetLevelRow(sheet, i);
                levelSheet.LevelRows.Add(lDataRow);
            }

            return levelSheet;
        }

        #endregion
    }
}