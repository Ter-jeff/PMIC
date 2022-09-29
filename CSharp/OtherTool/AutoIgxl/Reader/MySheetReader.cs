using Microsoft.Office.Interop.Excel;
using System;
using System.Text.RegularExpressions;

namespace AutoIgxl.Reader
{
    public class MySheetReader
    {
        protected int EndColNumber = -1;
        protected int EndRowNumber = -1;
        protected Worksheet ExcelWorksheet;

        protected int StartColNumber = -1;
        protected int StartRowNumber = -1;

        public bool IsLiked(string input, string patten)
        {
            if (patten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                patten.IndexOf(@".+", StringComparison.Ordinal) >= 0 ||
                patten.IndexOf(@"|", StringComparison.Ordinal) >= 0)
                return Regex.IsMatch(FormatCell(input), patten, RegexOptions.IgnoreCase);
            return FormatCell(input).Equals(patten, StringComparison.CurrentCultureIgnoreCase);
        }

        private string FormatCell(string text)
        {
            var result = text.Trim();

            result = ReplaceDoubleBlank(result);

            //result = result.Replace(" ", "_");

            result = result.ToUpper();

            return result;
        }

        private string ReplaceDoubleBlank(string text)
        {
            var result = text;
            do
            {
                result = result.Replace("  ", " ");
            } while (result.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return result;
        }

        public string GetMergedCellValue(Worksheet sheet, int rowNumber, int columnNumber)
        {
            Range range = sheet.Cells[rowNumber, columnNumber];
            return range.MergeCells
                ? GetCellValue(sheet, rowNumber, columnNumber)
                : GetCellValue(sheet, range.Row, range.Column);
        }

        public string GetCellValue(Worksheet sheet, int row, int column)
        {
            if (sheet == null) return "";
            if (row <= 0 || column <= 0) return "";
            //if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            //    return sheet.Cells[row, column].Formula;

            if (sheet.Cells[row, column].Value != null)
                return sheet.Cells[row, column].Value.ToString().Trim();
            if (sheet.Cells[row, column].Text != null)
                return sheet.Cells[row, column].Text.Trim();
            return "";
        }

        protected bool GetDimensions()
        {
            if (ExcelWorksheet.UsedRange != null)
            {
                StartColNumber = ExcelWorksheet.UsedRange.Column;
                StartRowNumber = ExcelWorksheet.UsedRange.Row;
                EndColNumber = ExcelWorksheet.UsedRange.Column + ExcelWorksheet.UsedRange.Columns.Count;
                EndRowNumber = ExcelWorksheet.UsedRange.Row + ExcelWorksheet.UsedRange.Rows.Count;
                return true;
            }

            return false;
        }
    }
}