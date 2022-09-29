using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.IO;

namespace IgxlData.IgxlReader
{
    public class ReadGlobalSpecSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public GlobalSpecSheet GetSheet(Stream stream, string sheetName)
        {
            var globalSpecSheet = new GlobalSpecSheet(sheetName, false);
            var isBackup = false;
            var i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var globalSpec = GetGlobalSpecRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(globalSpec.Symbol))
                        {
                            isBackup = true;
                            continue;
                        }

                        globalSpec.IsBackup = isBackup;
                        globalSpecSheet.AddRow(globalSpec);
                    }
                    i++;
                }
            }
            return globalSpecSheet;
        }

        private GlobalSpec GetGlobalSpecRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var globalSpec = new GlobalSpec();
            globalSpec.RowNum = row;
            globalSpec.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            globalSpec.ColumnA = content;
            content = GetCellText(arr, index);
            globalSpec.Symbol = content;
            index++;
            content = GetCellText(arr, index);
            globalSpec.Job = content;
            index++;
            content = GetCellText(arr, index);
            globalSpec.Value = content;
            index++;
            content = GetCellText(arr, index);
            globalSpec.Comment = content;
            return globalSpec;
        }

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
    }
}