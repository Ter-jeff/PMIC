using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadFlowSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public SubFlowSheet GetSheet(string fileName, bool isRemoveBackup = false)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName), isRemoveBackup);
        }

        public SubFlowSheet GetSheet(ExcelWorksheet sheet, bool isRemoveBackup = false)
        {
            SubFlowSheet subFlowSheet = new SubFlowSheet(sheet);
            int maxRowCount = sheet.Dimension.End.Row;

            // Set Row
            for (int i = StartRowIndex + 2; i <= maxRowCount; i++)
            {
                FlowRow flowRow = GetFlowRow(sheet, i);
                if (isRemoveBackup && string.IsNullOrEmpty(flowRow.OpCode))
                    break;
                subFlowSheet.AddRow(flowRow);
            }
            return subFlowSheet;
        }
        #endregion

        #region Private Function
        private FlowRow GetFlowRow(ExcelWorksheet sheet, int row)
        {
            FlowRow flowRow = new FlowRow();
            flowRow.SheetName = sheet.Name;
            int index = StartColumnIndex;
            flowRow.LineNum = row.ToString();
            var content = GetCellText(sheet, row, 1);
            flowRow.ColumnA = content;
            content = GetCellText(sheet, row, index);
            flowRow.Label = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Enable = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Job = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Part = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Env = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.OpCode = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Parameter = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Name = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Num = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.LoLim = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.HiLim = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Scale = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Units = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Format = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.BinPass = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.BinFail = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.SortPass = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.SortFail = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Result = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.PassAction = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.FailAction = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.State = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupSpecifier = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupSense = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupCondition = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.GroupName = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceSense = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceCondition = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DeviceName = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DebugAsume = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.DebugSites = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataElapsedTime = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataBackgroundType = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataSerialize = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataResourceLock = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.CtProfileDataFlowStepLocked = content;
            index++;
            content = GetCellText(sheet, row, index);
            flowRow.Comment = content;
            index++;
            var comment1List = new List<string>();
            for (int i = index; i <= sheet.Dimension.Columns; i++)
            {
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, index)))
                    comment1List.Add(GetCellText(sheet, row, index));
            }
            flowRow.Comment1 = string.Join("\t", comment1List);
            return flowRow;
        }
        #endregion
    }
}