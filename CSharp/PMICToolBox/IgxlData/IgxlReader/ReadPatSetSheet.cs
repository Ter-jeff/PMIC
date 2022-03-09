using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadPatSetSheet : IgxlSheetReader
    {
        private readonly List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public PatSetSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PatSetSheet GetSheet(Worksheet worksheet, bool isReadBackupRows = true)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet), isReadBackupRows);
        }

        public PatSetSheet GetSheet(ExcelWorksheet sheet, bool isReadBackupRows = true)
        {
            PatSetSheet patSetSheet = new PatSetSheet(sheet);
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
            List<PatSetRow> patSetRows = new List<PatSetRow>();
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                PatSetRow patSetRow = GetPatSetRow(sheet, i);
                if (!isReadBackupRows && string.IsNullOrEmpty(patSetRow.PatternSet))
                    break;

                patSetRows.Add(patSetRow);
            }

            var group = patSetRows.GroupBy(x => x.PatternSet).ToList();
            foreach (var item in group)
            {
                var patSet = new PatSet();
                patSet.PatSetName = item.Key;
                patSet.PatSetRows.AddRange(item);
                patSetSheet.AddPatSet(patSet);
            }
            return patSetSheet;
        }
        #endregion

        #region Private Function
        private PatSetRow GetPatSetRow(ExcelWorksheet sheet, int row)
        {
            PatSetRow patSetRow = new PatSetRow();
            for (int i = StartColumnIndex; i < _headList.Count; i++)
            {
                string lStrHead = _headList[i - StartColumnIndex];
                string lStrContent = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "PATTERN_SET":
                        patSetRow.PatternSet = lStrContent;
                        break;
                    case "TD_GROUP":
                        patSetRow.TdGroup = lStrContent;
                        break;
                    case "TIME_DOMAIN":
                        patSetRow.TimeDomain = lStrContent;
                        break;
                    case "ENABLE":
                        patSetRow.Enable = lStrContent;
                        break;
                    case "FILE/GROUP_NAME":
                        patSetRow.File = lStrContent;
                        break;
                    case "BURST":
                        patSetRow.Burst = lStrContent;
                        break;
                    case "START_LABEL":
                        patSetRow.StartLabel = lStrContent;
                        break;
                    case "STOP_LABEL":
                        patSetRow.StopLabel = lStrContent;
                        break;
                    case "COMMENT":
                        patSetRow.Comment = lStrContent;
                        break;
                }
            }
            return patSetRow;
        }
        #endregion
    }
}
