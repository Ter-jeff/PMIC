using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadLevelSheet : IgxlSheetReader
    {
        private List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public LevelSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public LevelSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public LevelSheet GetSheet(ExcelWorksheet sheet)
        {
            LevelSheet levelSheet = new LevelSheet(sheet);
            int maxRowCount = sheet.Dimension.End.Row;
            int maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                string lStrHead = lStrValue.Trim();
                _headList.Add(lStrHead);
            }

            // Set Row
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                LevelRow lDataRow = GetLevelRow(sheet, i);
                levelSheet.LevelRows.Add(lDataRow);
            }

            return levelSheet;
        }
        #endregion

        #region Private Function
        private LevelRow GetLevelRow(ExcelWorksheet sheet, int row)
        {
            string pinName = "";
            string parameter = "";
            string value = "";
            string comment = "";
            for (int i = StartColumnIndex; i <= sheet.Dimension.End.Column; i++)
            {
                string lStrHead = _headList[i - StartColumnIndex];
                string lStrContent = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "PIN/GROUP":
                        pinName = lStrContent;
                        break;
                    case "SEQ.":
                        break;
                    case "PARAMETER":
                        parameter = lStrContent;
                        break;
                    case "VALUE":
                        value = lStrContent.Replace("=", "");
                        break;
                    case "COMMENT":
                        comment = lStrContent;
                        break;
                }
            }
            return new LevelRow(pinName, parameter, value, comment, row);
        }
        #endregion
    }
}
