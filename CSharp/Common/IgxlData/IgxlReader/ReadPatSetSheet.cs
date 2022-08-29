using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadPatSetSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;
        private readonly List<string> _headers = new List<string>();

        #region Private Function

        private PatSetRow GetPatSetRow(ExcelWorksheet sheet, int row)
        {
            var patSetRow = new PatSetRow();
            for (var i = StartColumnIndex; i < _headers.Count; i++)
            {
                var lStrHead = _headers[i - StartColumnIndex];
                var content = GetCellText(sheet, row, i);
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

        #region public Function

        public PatSetSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PatSetSheet GetSheet(ExcelWorksheet sheet, bool isReadBackupRows = true)
        {
            var patSetSheet = new PatSetSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;
            var maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (var i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                var lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                var lStrHead = lStrValue.Trim();
                _headers.Add(lStrHead);
            }

            // Set Row
            var patSetRows = new List<PatSetRow>();
            for (var i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                var patSetRow = GetPatSetRow(sheet, i);
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
    }
}