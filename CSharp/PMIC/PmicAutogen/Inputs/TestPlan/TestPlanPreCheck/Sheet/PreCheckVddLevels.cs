using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;
using System;

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
            foreach (var sheetConfig in SheetConfigs)
            {
                if (sheetConfig.Type == EnumColumn.None)
                    continue;
                var columnIndex = GetColumnIndex(sheetConfig);
                if (columnIndex != -1)
                    for (var i = StartRow + 1; i <= _excelWorksheet.Dimension.End.Row; i++)
                    {
                        string errorMessage;
                        var value = _excelWorksheet.GetCellValue(i, columnIndex);
                        if (sheetConfig.HeaderName.Equals("SEQ", StringComparison.CurrentCultureIgnoreCase) &&
                            value.Equals("x", StringComparison.CurrentCultureIgnoreCase))
                            continue;
                        if (!SheetStructureManager.JudgeCell(sheetConfig.Type, value, out errorMessage))
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, i,
                                columnIndex, errorMessage);
                    }
            }
        }
    }
}