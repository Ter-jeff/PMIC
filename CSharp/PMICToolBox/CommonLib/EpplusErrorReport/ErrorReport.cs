﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using CommonLib.Utility;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace CommonLib.EpplusErrorReport
{
    public class ErrorReport
    {
        private readonly List<Error> _errors;

        private string ReportName
        {
            get
            {
                if (_errors.Count == 0)
                    return "";

                return _errors[0].ErrorType + "Report";
            }
        }

        public ErrorReport(List<Error> errors)
        {
            _errors = errors;
        }

        private List<string> GetErrorSubType()
        {
            return _errors.GroupBy(p => p.ErrorType).Select(p => p.Key).ToList();
        }

        private List<Error> GetErrorsBySubType(object subtype)
        {
            return _errors.Where(p => p.ErrorType.Equals(subtype)).ToList();
        }

        public void WriteReport(List<ExcelWorkbook> workbooks, string errorReprortName = "", string summaryReport = "SummaryReport")
        {
            if (_errors.Count == 0)
                return;

            var finalErrorReportName = !string.IsNullOrEmpty(errorReprortName) ? errorReprortName : ReportName;

            try
            {
                List<Error> errorsHaveWritten = new List<Error>();
                for (int i = 0; i < workbooks.Count - 1; i++)
                {
                    List<string> names = workbooks[i].Worksheets.Select(x => x.Name).ToList();
                    List<Error> errorsNeedToWrite = _errors.Where(x =>
                            names.Exists(y => y.Equals(x.SheetName, StringComparison.CurrentCultureIgnoreCase)))
                        .ToList();
                    WriteErrors(workbooks[i], errorsNeedToWrite, finalErrorReportName, summaryReport);
                    errorsHaveWritten.AddRange(errorsNeedToWrite);
                    workbooks[i].Worksheets.First().Select();
                }

                WriteErrors(workbooks.Last(), _errors.Where(p => !errorsHaveWritten.Exists(a => a.SheetName == p.SheetName)).ToList(), finalErrorReportName, summaryReport);
                workbooks.Last().Worksheets.First().Select();
            }
            catch (Exception e)
            {
                throw new Exception("Write General ErrorReport failed for " + finalErrorReportName + "  " + e.Message);
            }
        }

        private void WriteErrors(ExcelWorkbook workbook, List<Error> errors, string reportName, string summaryReport = "SummaryReport")
        {
            if (errors.Count == 0)
                return;

            workbook.DeleteSheet(reportName);

            ExcelWorksheet workSheet = workbook.AddSheet(reportName);
            IExcelConditionalFormattingEqual condition = workSheet.ConditionalFormatting.AddEqual(new ExcelAddress("$B:$B"));
            condition.Style.Fill.PatternType = ExcelFillStyle.Solid;
            condition.Style.Fill.BackgroundColor.Color = Color.Red;
            condition.Formula = "\"Error\"";

            object[] headers = { "ErrorType", "Level", "Link", "SheetName", "Row", "Col", "ErrorMessage", "Count" };
            workSheet.Cells[1, 1].LoadFromArrays(new List<object[]> { headers }).Style.Font.Bold = true;

            int currentRow = 2;
            int count = 0;
            List<string> typeList = GetErrorSubType();
            List<object[]> content = new List<object[]>();
            foreach (object errorType in typeList)
            {
                List<Error> subErrors = GetErrorsBySubType(errorType);
                foreach (Error error in subErrors)
                {
                    object[] row = new object[7 + error.Comments.Count];
                    row[0] = error.ErrorType;
                    row[1] = error.ErrorLevel.ToString();
                    if (workbook.Worksheets.Any(x =>
                            x.Name.Equals(error.SheetName, StringComparison.OrdinalIgnoreCase)) && error.RowNum > 0)
                    {
                        row[2] = error.Link;
                        int sameCell = subErrors
                          .Where(x => x.SheetName == error.SheetName && x.RowNum == error.RowNum &&
                                      x.ColNum == error.ColNum).Max(y => (int)y.ErrorLevel);
                        ErrorLevel errorLevel = sameCell == 2 ? ErrorLevel.Error : ErrorLevel.Warning;
                        SetErrorColor(workbook, error, errorLevel);
                    }

                    row[3] = error.SheetName;
                    row[4] = error.RowNum;
                    row[5] = error.ColNum;
                    row[6] = error.Message;
                    for (int index = 0; index < error.Comments.Count; index++)
                    {
                        row[index + 7] = error.Comments[index];
                    }

                    content.Add(row);
                    currentRow++;
                }

                for (int i = currentRow - subErrors.Count; i < currentRow; i++)
                {
                    workSheet.Row(i).OutlineLevel = 1;
                    workSheet.Row(i).Collapsed = true;
                }
                object[] subSumRow = { errorType.ToString(), "", "", "", "", "", "", subErrors.Count };
                content.Add(subSumRow);
                count += subErrors.Count;
                currentRow++;
            }

            object[] totalRow = { "Total error", "", "", "", "", "", "", count };
            content.Add(totalRow);
            workSheet.Cells[currentRow, 1, currentRow, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[currentRow, 1, currentRow, 8].Style.Fill.BackgroundColor.SetColor(Color.Red);
            workSheet.Cells[2, 1].LoadFromArrays(content);
            workSheet.SetFormula(3);
            workSheet.Cells["1:1"].AutoFilter = true;
            workSheet.Cells.AutoFitColumns();
            workSheet.Column(7).Width = 100;

            if (typeList.Count > 1 && !string.IsNullOrEmpty(summaryReport))
                AppendSummaryReport(workbook, reportName, summaryReport, count);
        }

        private void SetErrorColor(ExcelWorkbook workbook, Error error, ErrorLevel errorLevel)
        {
            ExcelWorksheet workSheet = workbook.Worksheets[error.SheetName];
            ExcelRange range = error.ColNum > 0
                ? workSheet.Cells[error.RowNum, error.ColNum, error.RowNum, error.ColNum]
                : workSheet.Cells[error.RowNum, 1, error.RowNum, workSheet.Dimension.End.Column];

            if (errorLevel == ErrorLevel.Error)
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.Red);
            }
            else if (error.ErrorLevel == ErrorLevel.Warning)
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            }
        }

        private void AppendSummaryReport(ExcelWorkbook workbook, string reportName, string summaryReport, int count)
        {
            ExcelWorksheet sheet;
            if (workbook.Worksheets.Any(x => x.Name.Equals(summaryReport, StringComparison.OrdinalIgnoreCase)))
            {
                sheet = workbook.Worksheets[summaryReport];
            }
            else
            {
                sheet = workbook.AddSheet(summaryReport);
                object[] headers = { "ReportType", "Count", "Link" };
                sheet.Cells[1, 1].LoadFromArrays(new List<object[]> { headers }).Style.Font.Bold = true;
            }

            int startRow = 1;
            if (sheet.Dimension!=null &&　sheet.Dimension.End != null)
                startRow = sheet.Dimension.End.Row + 1;
            sheet.Cells[startRow, 1].Value = reportName;
            sheet.Cells[startRow, 2].Value = count;
            sheet.Cells[startRow, 3].Value = "Link";
            sheet.Cells[startRow, 3].Hyperlink = new Uri("#" + "'" + reportName + "'" + "!A1", UriKind.Relative);
            sheet.Cells[startRow, 3].Style.Font.UnderLine = true;
            sheet.Cells[startRow, 3].Style.Font.Color.SetColor(Color.Blue);
            sheet.Cells.AutoFitColumns();
        }
    }
}