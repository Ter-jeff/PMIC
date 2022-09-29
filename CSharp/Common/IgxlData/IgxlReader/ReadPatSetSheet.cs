using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadPatSetSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;
        private readonly List<string> _headers = new List<string>();

        public PatSetSheet GetSheet(Stream stream, string sheetName)
        {
            var patSetSheet = new PatSetSheet(sheetName);
            var isBackup = false;
            var i = 1;
            var patSetRows = new List<PatSetRow>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var patSetRow = GetPatSetRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(patSetRow.PatternSet))
                        {
                            isBackup = true;
                            continue;
                        }
                        patSetRow.IsBackup = isBackup;
                        patSetRows.Add(patSetRow);
                    }
                    i++;
                }

                var group = patSetRows.ChunkBy(x => x.PatternSet).ToList();
                foreach (var item in group)
                {
                    var patSet = new PatSet();
                    patSet.PatSetName = item.Key;
                    patSet.PatSetRows.AddRange(item);
                    patSet.IsBackup = item.All(x => x.IsBackup);
                    patSetSheet.AddPatSet(patSet);
                }
            }
            return patSetSheet;
        }

        private PatSetRow GetPatSetRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var patSetRow = new PatSetRow();
            patSetRow.RowNum = row;
            patSetRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            patSetRow.ColumnA = content;
            content = GetCellText(arr, index);
            patSetRow.PatternSet = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.TimeDomain = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.Enable = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.File = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.Burst = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.StartLabel = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.StopLabel = content;
            index++;
            content = GetCellText(arr, index);
            patSetRow.Comment = content;
            return patSetRow;
        }

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

            var group = patSetRows.ChunkBy(x => x.PatternSet).ToList();
            foreach (var item in group)
            {
                var patSet = new PatSet();
                patSet.PatSetName = item.Key;
                patSet.PatSetRows.AddRange(item);
                patSet.IsBackup = item.All(x => x.IsBackup);
                patSetSheet.AddPatSet(patSet);
            }

            return patSetSheet;
        }
    }
}