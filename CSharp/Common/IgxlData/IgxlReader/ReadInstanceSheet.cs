using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.IO;

namespace IgxlData.IgxlReader
{
    public class ReadInstanceSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public InstanceSheet GetSheet(Stream stream, string sheetName)
        {
            var instanceSheet = new InstanceSheet(sheetName);
            var isBackup = false;
            var i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var instanceRow = GetInstanceRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(instanceRow.TestName))
                        {
                            isBackup = true;
                            continue;
                        }

                        instanceRow.IsBackup = isBackup;
                        instanceSheet.AddRow(instanceRow);
                    }
                    i++;
                }
            }
            return instanceSheet;
        }

        public InstanceSheet GetSheet(string fileName)
        {
            var sheetName = Path.GetFileNameWithoutExtension(fileName);
            var lines = File.ReadAllLines(fileName);
            var instanceSheet = new InstanceSheet(sheetName);
            var maxRowCount = lines.Length;
            var backup = false;
            for (var i = StartRowIndex + 1; i < maxRowCount; i++)
            {
                var instanceRow = GetInstanceRow(lines[i], sheetName, i);
                if (string.IsNullOrEmpty(instanceRow.TestName))
                    backup = true;
                instanceRow.IsBackup = backup;
                instanceSheet.AddRow(instanceRow);
            }

            return instanceSheet;
        }

        public InstanceSheet GetSheet(ExcelWorksheet sheet)
        {
            var instanceSheet = new InstanceSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;
            var backup = false;
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var instanceRow = GetInstanceRow(sheet, i);
                if (string.IsNullOrEmpty(instanceRow.TestName))
                    backup = true;
                instanceRow.IsBackup = backup;
                instanceSheet.AddRow(instanceRow);
            }

            return instanceSheet;
        }

        private InstanceRow GetInstanceRow(ExcelWorksheet sheet, int row)
        {
            var instanceRow = new InstanceRow();
            instanceRow.RowNum = row;
            var index = StartColumnIndex - 1;
            var content = GetCellText(sheet, row, 1);
            instanceRow.ColumnA = content;
            content = GetCellText(sheet, row, index);
            instanceRow.TestName = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.Type = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.Name = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.CalledAs = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.DcCategory = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.DcSelector = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.AcCategory = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.AcSelector = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.TimeSets = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.EdgeSets = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.PinLevels = content;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.MixedSignalTiming = content;
            index++;
            instanceRow.Overlay = string.Empty;
            index++;
            content = GetCellText(sheet, row, index);
            instanceRow.ArgList = content;
            for (var i = 1; i <= 130; i++)
            {
                content = GetCellText(sheet, row, index + i);
                instanceRow.Args.Add(content);
            }

            index += 130;
            content = GetCellText(sheet, row, index);
            instanceRow.Comment = content;
            return instanceRow;
        }

        private InstanceRow GetInstanceRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var instanceRow = new InstanceRow();
            instanceRow.RowNum = row;
            instanceRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            instanceRow.ColumnA = content;
            content = GetCellText(arr, index);
            instanceRow.TestName = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.Type = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.Name = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.CalledAs = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.DcCategory = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.DcSelector = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.AcCategory = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.AcSelector = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.TimeSets = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.EdgeSets = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.PinLevels = content;
            index++;
            content = GetCellText(arr, index);
            instanceRow.MixedSignalTiming = content;
            index++;
            instanceRow.Overlay = string.Empty;
            index++;
            content = GetCellText(arr, index);
            instanceRow.ArgList = content;
            for (var i = 1; i <= 130; i++)
            {
                content = GetCellText(arr, index + i);
                instanceRow.Args.Add(content);
            }

            index += 130;
            content = GetCellText(arr, index);
            instanceRow.Comment = content;
            return instanceRow;
        }
    }
}