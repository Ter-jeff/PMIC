using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadInstanceSheet : IgxlSheetReader
    {
        private readonly List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public InstanceSheet GetSheet(Worksheet worksheet, bool isRemoveBackup = false)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public InstanceSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public InstanceSheet GetSheet(ExcelWorksheet sheet, bool isReadBackupRows = true)
        {
            var instanceSheet = new InstanceSheet(sheet);
            int maxRowCount = sheet.Dimension.End.Row;
            int maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                string lStrValue2 = GetCellText(sheet, StartRowIndex + 1, i);
                string lStrHead = lStrValue.Trim() + "_" + lStrValue2.Trim();
                _headList.Add(lStrHead);
            }

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
            int index = 2;
            string lStrContent;
            lStrContent = GetCellText(sheet, row, 1);
            instanceRow.ColumnA = lStrContent;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.TestName = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.Type = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.Name = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.CalledAs = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.DcCategory = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.DcSelector = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.AcCategory = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.AcSelector = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.TimeSets = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.EdgeSets = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.PinLevels = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.MixedSignalTiming = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.Overlay = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.ArgList = lStrContent;
            for (int i = 1; i <= 130; i++)
            {
                lStrContent = GetCellText(sheet, row, index + i);
                instanceRow.Args.Add(lStrContent);
            }
            index += 130;
            lStrContent = GetCellText(sheet, row, index);
            instanceRow.Comment = lStrContent;
            return instanceRow;
        }
        #endregion
    }
}
