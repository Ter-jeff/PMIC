using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace IgxlData
{
    public class ReadPatSubroutineSheet : IgxlSheetReader
    {
        private List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public PatSetSubSheet GetSheet(Stream stream, string sheetName)
        {
            var patSetSubSheet = new PatSetSubSheet(sheetName);
            var isBackup = false;
            var i = 1;
            var patSetSubRows = new List<PatSetSubRow>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var patSetRow = GetPatSetSubRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(patSetRow.PatternFileName))
                        {
                            isBackup = true;
                            continue;
                        }
                        patSetRow.IsBackup = isBackup;
                        patSetSubRows.Add(patSetRow);
                    }
                    i++;
                }
            }
            patSetSubSheet.PatSetSubRows = patSetSubRows;
            return patSetSubSheet;
        }

        private PatSetSubRow GetPatSetSubRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var patSetSubRow = new PatSetSubRow();
            patSetSubRow.RowNum = row;
            patSetSubRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            patSetSubRow.ColumnA = content;
            content = GetCellText(arr, index);
            patSetSubRow.PatternFileName = content;
            index++;
            content = GetCellText(arr, index);
            patSetSubRow.Comment = content;
            return patSetSubRow;
        }

        public PatSetSubSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PatSetSubSheet GetSheet(ExcelWorksheet sheet)
        {
            _headList = new List<string>();
            var patSetSubSheet = new PatSetSubSheet(sheet);
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
            var patSetSuRows = new List<PatSetSubRow>();
            var isBackup = false;
            for (var i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                var patSetSubRow = GetPatSetSubRow(sheet, i);
                if (string.IsNullOrEmpty(patSetSubRow.PatternFileName))
                {
                    isBackup = true;
                    continue;
                }
                if (isBackup)
                    patSetSubRow.IsBackup = true;

                patSetSuRows.Add(patSetSubRow);
            }
            patSetSubSheet.PatSetSubRows = patSetSuRows;

            return patSetSubSheet;
        }

        private PatSetSubRow GetPatSetSubRow(ExcelWorksheet sheet, int row)
        {
            var patSetSubRow = new PatSetSubRow();
            for (var i = StartColumnIndex; i < StartColumnIndex + _headList.Count; i++)
            {
                var lStrHead = _headList[i - StartColumnIndex];
                var lStrContent = GetCellText(sheet, row, i);
                patSetSubRow.SheetName = sheet.Name;
                patSetSubRow.RowNum = row;
                switch (FormatStringForCompare(lStrHead))
                {
                    case "PATTERN_FILENAME":
                        patSetSubRow.PatternFileName = lStrContent;
                        break;
                    case "COMMENT":
                        patSetSubRow.Comment = lStrContent;
                        break;
                }
            }
            return patSetSubRow;
        }
    }
}
