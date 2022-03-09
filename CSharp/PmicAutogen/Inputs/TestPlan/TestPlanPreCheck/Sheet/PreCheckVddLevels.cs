using AutomationCommon.EpplusErrorReport;
using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckVddLevels : PreCheckBase
    {
        public PreCheckVddLevels(ExcelWorkbook testPlanWorkbook, string sheetName) : base(testPlanWorkbook, sheetName)
        {
        }

        protected override bool CheckBusiness()
        {
            return true;
        }

        protected override void CheckFormat()
        {
            foreach (var sheetConfig in _sheetConfigs)
            {
                if (sheetConfig.Type == EnumColumn.None)
                    continue;
                var columnIndex = GetColumnIndex(sheetConfig);
                if (columnIndex != -1)
                    for (var i = StartRow + 1; i <= Worksheet.Dimension.End.Row; i++)
                    {
                        string errorMessage;
                        var value = EpplusOperation.GetCellValue(Worksheet, i, columnIndex);
                        if (sheetConfig.HeaderName.Equals("SEQ", System.StringComparison.CurrentCultureIgnoreCase) &&
                            value.Equals("x",System.StringComparison.CurrentCultureIgnoreCase))
                        {
                            continue;
                        }
                        if (!SheetStructureManager.JudgeCell(sheetConfig.Type, value, out errorMessage))
                            EpplusErrorManager.AddError(BasicErrorType.FormatError, ErrorLevel.Error, SheetName, i,
                                columnIndex, errorMessage);
                    }
            }
        }
    }
}