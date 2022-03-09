using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadPatSetSheet : IgxlSheetReader
    {
        private readonly List<string> _headers = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public PatSetSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
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
                _headers.Add(lStrHead);
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
            for (int i = StartColumnIndex; i < _headers.Count; i++)
            {
                string lStrHead = _headers[i - StartColumnIndex];
                string content = GetCellText(sheet, row, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "PATTERN_SET":
                        patSetRow.PatternSet = content;
                        break;
                    case "TD_GROUP":
                        patSetRow.TdGroup = content;
                        break;
                    case "TIME_DOMAIN":
                        patSetRow.TimeDomain = content;
                        break;
                    case "ENABLE":
                        patSetRow.Enable = content;
                        break;
                    case "FILE/GROUP_NAME":
                        patSetRow.File = content;
                        break;
                    case "BURST":
                        patSetRow.Burst = content;
                        break;
                    case "START_LABEL":
                        patSetRow.StartLabel = content;
                        break;
                    case "STOP_LABEL":
                        patSetRow.StopLabel = content;
                        break;
                    case "COMMENT":
                        patSetRow.Comment = content;
                        break;
                }
            }
            return patSetRow;
        }
        #endregion
    }
}
