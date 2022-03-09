using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadInstanceSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public InstanceSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public InstanceSheet GetSheet(ExcelWorksheet sheet, bool isReadBackupRows = true)
        {
            InstanceSheet instanceSheet = new InstanceSheet(sheet);
            int maxRowCount = sheet.Dimension.End.Row;

            // Set Row
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                InstanceRow instanceRow = GetInstanceRow(sheet, i);
                if (!isReadBackupRows && string.IsNullOrEmpty(instanceRow.TestName))
                    break;
                instanceSheet.AddRow(instanceRow);
            }

            return instanceSheet;
        }
        #endregion

        #region Private Function
        private InstanceRow GetInstanceRow(ExcelWorksheet sheet, int row)
        {
            InstanceRow instanceRow = new InstanceRow();
            instanceRow.RowNum = row;
            int index = StartColumnIndex;
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
            for (int i = 1; i <= 130; i++)
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
    }
}
