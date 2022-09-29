using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;
using System;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckDcTestContinuity : PreCheckBase
    {
        public PreCheckDcTestContinuity(ExcelWorkbook testPlanWorkbook, string sheetName) : base(testPlanWorkbook,
            sheetName)
        {
        }

        protected override bool CheckBusiness()
        {
            var result = true;

            for (var j = StartColumn; j <= _excelWorksheet.Dimension.End.Column; j++)
            {
                var header = _excelWorksheet.GetMergedCellValue(StartRow, j);
                if (IsLiked(header, DcTestContiRow.ConHeaderCondition))
                    if (!CheckConditionFormat(j))
                        result = false;

                if (IsLiked(header, DcTestContiRow.ConHeaderLimit))
                    if (!CheckLimitFormat(j))
                        result = false;
            }

            return result;
        }

        private bool CheckConditionFormat(int column)
        {
            var result = true;
            const string pattern = @"^\w*=\s*[-+]?\d*[.]?\d*\w*$";
            for (var i = StartRow + 2; i <= _excelWorksheet.Dimension.End.Row; i++)
            {
                var value = _excelWorksheet.GetCellValue(i, column);
                if (value == "")
                    return true;

                if (Regex.IsMatch(value, pattern) == false)
                {
                    result = false;
                    var errorMessage = string.Format("Column[{0}] Can not recognize the condition {1}",
                        DcTestContiRow.ConHeaderCondition, value);
                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, i, column,
                        errorMessage);
                }

                if (!Regex.IsMatch(value, "Isource=.*|Isink=.*", RegexOptions.IgnoreCase))
                {
                    var errorMessage = string.Format("Column[{0}] Can not recognize the condition {1}",
                        DcTestContiRow.ConHeaderCondition, value);
                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, i, column,
                        errorMessage);
                }
            }

            return result;
        }

        private bool CheckLimitFormat(int column)
        {
            var result = true;

            var mergeCellCnt = 2;
            var range = _excelWorksheet.MergedCells[StartRow, 1];
            if (range != null)
            {
                var address = new ExcelAddress(range);
                mergeCellCnt = address.End.Row - address.Start.Row;
            }

            for (var i = StartRow + mergeCellCnt + 1; i < _excelWorksheet.Dimension.End.Row; i++)
            {
                var value = _excelWorksheet.GetCellValue(i, column);
                if (value == "")
                    return true;
                double limit;

                if (!double.TryParse(value, out limit))
                    if (!(value.EndsWith("V", StringComparison.CurrentCultureIgnoreCase) ||
                          value.EndsWith("I", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("The format of {0} is not allowed !!!", value);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, i, column,
                            errorMessage);
                        result = false;
                    }
            }

            return result;
        }
    }
}