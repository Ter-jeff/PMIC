using CommonLib.Enum;
using CommonLib.Extension;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CommonLib.ErrorReport
{
    public class ErrorInstance
    {
        public static ErrorInstance Instance = new ErrorInstance();
        private readonly List<Error> _errors;

        private ErrorInstance()
        {
            _errors = new List<Error>();
        }

        public List<Error> GetErrors()
        {
            return _errors;
        }

        public void AddError(Error error)
        {
            _errors.Add(error);
        }

        public void AddErrors(List<Error> errors)
        {
            _errors.AddRange(errors);
        }

        public void Initialize()
        {
            _errors.Clear();
        }

        public int GetErrorCount()
        {
            return _errors.Count;
        }

        public int GetErrorCountByType(string type)
        {
            return GetErrorsByErrorType(type).Count;
        }

        private List<Error> GetErrorsByErrorType(string type)
        {
            return _errors.Where(a => a.EnumErrorType.ToString().Equals(type, StringComparison.CurrentCulture)).ToList();
        }

        private List<string> GetEnumErrorTypes()
        {
            return _errors.GroupBy(p => p.EnumErrorType.ToString()).Select(p => p.Key).ToList();
        }

        #region Report
        public void WriteErrors(Worksheet worksheet)
        {
            worksheet.UsedRange.ClearOutline();
            Workbook workbook = worksheet.Parent;
            FormatCondition formatCondition = worksheet.Columns[2].FormatConditions.Add(XlFormatConditionType.xlCellValue,
               XlFormatConditionOperator.xlEqual, "Error", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            formatCondition.Interior.Color = Color.Red;

            string[] headers = { "ErrorType", "Level", "SheetName", "Link", "Row", "Col", "Message" };
            Range range = worksheet.Cells[1, 1];
            range.LoadFromArrays(new List<string[]> { headers }).Font.Bold = true;

            var currentRow = 2;
            var count = 0;
            var typeList = GetEnumErrorTypes();
            foreach (var errorType in typeList)
            {
                var subErrors = GetErrorsByErrorType(errorType);
                Range subRange = worksheet.Cells[currentRow, 1];
                subRange.LoadFromCollection(subErrors, false).Font.Bold = false;
                var formula = "=COUNTA(" + worksheet.Cells[currentRow, 7].Address + ":" +
                            worksheet.Cells[currentRow + subErrors.Count() - 1, 7].Address + ")";
                currentRow += subErrors.Count();
                foreach (var error in subErrors)
                {
                    if (workbook.IsSheetExist(error.SheetName) && error.RowNum > 0)
                    {
                        var sameCell = subErrors
                            .Where(x => x.SheetName == error.SheetName && x.RowNum == error.RowNum &&
                                        x.ColNum == error.ColNum).Max(y => (int)y.ErrorLevel);
                        var errorLevel = sameCell == 1 ? EnumErrorLevel.Error : EnumErrorLevel.Warning;
                        SetErrorColor(workbook, error, errorLevel);
                    }
                }
                object[] subSumRow = { errorType.ToString(), "", "", "", "", "", formula };
                subRange = worksheet.Cells[currentRow, 1];
                subRange.LoadFromArrays(new List<object[]> { subSumRow });
                count += subErrors.Count;
                currentRow++;
            }

            var sumformula = "=SUM(" + worksheet.Cells[2, 7].Address + ":" +
                           worksheet.Cells[currentRow - 1, 7].Address + ")";
            object[] totalRow = { "Total error", "", "", "", "", "", sumformula };
            Range rangeTotalRow = worksheet.Cells[currentRow, 1];
            rangeTotalRow.LoadFromArrays(new List<object[]> { totalRow }).Interior.Color = ColorTranslator.ToOle(Color.Yellow);

            worksheet.UsedRange.AutoFilter(1);
            worksheet.Columns.AutoFit();
            worksheet.Columns[7].ColumnWidth = 70;
            worksheet.UsedRange.AutoOutline();
            //worksheet.Outline.ShowLevels(2);
        }

        private void SetErrorColor(Workbook workbook, Error error, EnumErrorLevel errorLevel)
        {
            Worksheet workSheet = workbook.Worksheets[error.SheetName];
            Range range = error.ColNum > 0
                ? workSheet.GetRange(error.RowNum, error.ColNum, error.RowNum, error.ColNum)
                : workSheet.GetRange(error.RowNum, 1, error.RowNum,
                workSheet.UsedRange.Row + workSheet.UsedRange.Rows.Count);

            if (errorLevel == EnumErrorLevel.Error)
            {
                range.Interior.Color = ColorTranslator.ToOle(Color.Pink);
                range.Font.Color = ColorTranslator.ToOle(Color.Red);
            }
            else if (error.ErrorLevel == EnumErrorLevel.Warning)
            {
                range.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                range.Font.Color = ColorTranslator.ToOle(Color.Red);
            }
        }
        #endregion
    }
}