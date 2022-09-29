using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CommonReaderLib
{
    public class MySheetReader
    {
        protected ExcelWorksheet ExcelWorksheet;

        protected int StartRowNumber = -1;
        protected int EndRowNumber = -1;
        protected int StartColNumber = -1;
        protected int EndColNumber = -1;

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

        protected bool GetDimensions()
        {
            if (ExcelWorksheet.Dimension != null)
            {
                StartColNumber = ExcelWorksheet.Dimension.Start.Column;
                StartRowNumber = ExcelWorksheet.Dimension.Start.Row;
                EndColNumber = ExcelWorksheet.Dimension.End.Column;
                EndRowNumber = ExcelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        public ExcelWorksheet ConvertCsvToExcelSheet(string fileName)
        {
            var excelPackage = new ExcelPackage();
            var sheetName = Path.GetFileNameWithoutExtension(fileName);
            var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            var index = 0;
            using (var sr = new StreamReader(fileName))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    index++;
                    if (line != null)
                    {
                        var arr = line.Split(new[] { ',' }, StringSplitOptions.None);
                        var cnt = 0;
                        foreach (var item in arr)
                        {
                            sheet.Cells[index, 1 + cnt].Value = item;
                            cnt++;
                        }
                    }
                }
            }

            return sheet;
        }
    }
}