using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadFlowSheet : IgxlSheetReader
    {
        private List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function

        public SubFlowSheet GetSheet(Worksheet worksheet, bool isRemoveBackup = false)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet), isRemoveBackup);
        }

        public SubFlowSheet GetSheet(string fileName, bool isRemoveBackup = false)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName), isRemoveBackup);
        }

        public SubFlowSheet GetSheet(ExcelWorksheet sheet, bool isRemoveBackup = false)
        {
            SubFlowSheet subFlowSheet = new SubFlowSheet(sheet);
            if (sheet.Dimension == null) return subFlowSheet;

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
                FlowRow flowRow = GetFlowRow(sheet, i);
                if (isRemoveBackup && string.IsNullOrEmpty(flowRow.Opcode))
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
            int index = 2;
            string lStrContent;
            flowRow.LineNum = row.ToString();
            lStrContent = GetCellText(sheet, row, 1);
            flowRow.ColumnA = lStrContent;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Label = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Enable = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Job = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Part = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Env = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Opcode = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Parameter = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.TName = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.TNum = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.LoLim = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.HiLim = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Scale = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Units = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Format = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.BinPass = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.BinFail = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.SortPass = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.SortFail = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Result = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.PassAction = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.FailAction = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.State = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.GroupSpecifier = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.GroupSense = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.GroupCondition = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.GroupName = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.DeviceSense = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.DeviceCondition = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.DeviceName = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.DebugAsume = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.DebugSites = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.CtProfileDataElapsedTime = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.CtProfileDataBackgroundType = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.CtProfileDataSerialize = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.CtProfileDataResourceLock = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.CtProfileDataFlowStepLocked = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            flowRow.Comment = lStrContent;
            index++;
            var comment1List = new List<string>();
            for (int i = index; i <= sheet.Dimension.Columns; i++)
            {
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, i)))
                    comment1List.Add(GetCellText(sheet, row, i));
            }
            flowRow.Comment1 = string.Join("\t", comment1List);
            return flowRow;
        }
        #endregion
    }
}