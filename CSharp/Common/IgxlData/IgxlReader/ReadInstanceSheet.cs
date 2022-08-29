using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadInstanceSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region Private Function

        private InstanceRow GetInstanceRow(ExcelWorksheet sheet, int row)
        {
            var instanceRow = new InstanceRow();
            instanceRow.RowNum = row;
            var index = StartColumnIndex;
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

        #endregion

        #region public Function

        public InstanceSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public InstanceSheet GetSheet(ExcelWorksheet sheet, bool isReadBackupRows = true)
        {
            var instanceSheet = new InstanceSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;

            // Set Row
            for (var i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                var instanceRow = GetInstanceRow(sheet, i);
                if (!isReadBackupRows && string.IsNullOrEmpty(instanceRow.TestName))
                    break;
                instanceSheet.AddRow(instanceRow);
            }

            return instanceSheet;
        }

        #endregion
    }
}