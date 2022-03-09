using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace CommonLib.Utility
{
    public class EpplusOperation
    {
        public static bool IsOpened(string filePath)
        {
            if (!File.Exists(filePath)) return false;
            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        public static Dictionary<string, int> GetHeaderOrder(ExcelWorksheet sheet, int startRow = 1)
        {
            Dictionary<string, int> headerOrder = new Dictionary<string, int>();
            if (sheet.Dimension == null)
                return headerOrder;
            int endCol = sheet.Dimension.End.Column;
            for (int i = 1; i <= endCol; i++)
            {
                if (sheet.Cells[startRow, i].Value != null)
                {
                    string header = sheet.Cells[startRow, i].Value.ToString().Trim();
                    if (!headerOrder.ContainsKey(header))
                        headerOrder.Add(header, i);
                }
            }
            return headerOrder;
        }

        public static string FloorMinValue(string value, string pwrSupplyRes)
        {
            if (Regex.IsMatch(value, @"\=.*\*", RegexOptions.IgnoreCase))
            {
                string resultValue = value.Replace("=", "");
                double outD;
                Double.TryParse(pwrSupplyRes, out outD);
                if (outD > 0.001)
                    resultValue = "=FLOOR(" + resultValue + "," + outD + ")";
                else
                    resultValue = "=" + resultValue;
                return resultValue;
            }
            return value;
        }

        public static string CeilingMaxValue(string value, string pwrSupplyRes)
        {
            if (Regex.IsMatch(value, @"\=.*\*", RegexOptions.IgnoreCase))
            {
                string resultValue = value.Replace("=", "");
                double outD;
                Double.TryParse(pwrSupplyRes, out outD);
                if (outD > 0.001)
                    resultValue = "=CEILING(" + resultValue + "," + outD + ")";
                else
                    resultValue = "=" + resultValue;
                return resultValue;
            }
            return value;
        }

        public static string GetMergedCellValue(ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null ?
                GetCellValue(sheet, rowNumber, columnNumber) :
                GetCellValue(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column));
        }

        public static string GetCellFormula(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null) return "";
            if (row <= 0 || column <= 0) return "";
            string value = sheet.Cells[row, column].Formula;
            if (value.Equals(""))
            {
                return GetCellValue(sheet, row, column);
            }
            else
            {
                return value;
            }
        }

        public static string GetCellText(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null) return "";
            if (row <= 0 || column <= 0) return "";
            string value = sheet.Cells[row, column].Text;
            if (value != null)
            {
                if (value.Contains("%"))
                    return value.Trim();
                return GetCellValue(sheet, row, column);
            }
            return "";
        }

        public static string GetCellValue(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null) return "";
            if (row <= 0 || column <= 0) return "";
            //if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            //    return sheet.Cells[row, column].Formula;

            if (sheet.Cells[row, column].Value != null)
                return sheet.Cells[row, column].Value.ToString();
            if (sheet.Cells[row, column].Text != null)
                return sheet.Cells[row, column].Text;
            return "";
        }

        public static string GetCellValueOld(ExcelWorksheet wSheet, int row, int column)
        {
            if (row <= 0 || column <= 0) return "";
            if (wSheet.Cells[row, column].Value != null)
                return wSheet.Cells[row, column].Value.ToString().Trim();
            return "";
        }

        private static string ReplaceDoubleBlank(string text)
        {
            string result = text;
            do
            {
                result = result.Replace("  ", " ");
            } while (result.IndexOf("  ", StringComparison.Ordinal) >= 0);
            return result;
        }

        private static string FormatCell(string text)
        {
            string result = text.Trim();

            result = ReplaceDoubleBlank(result);

            result = result.Replace(" ", "_");

            result = result.ToUpper();

            return result;
        }

        public static bool IsLiked(string pStrInput, string pStrPatten)
        {
            if (pStrPatten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@".+", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@"|", StringComparison.Ordinal) >= 0)
            {
                bool value = Regex.IsMatch(FormatCell(pStrInput), FormatCell(pStrPatten));
                return value;
            }
            else
            {
                bool value = FormatCell(pStrInput) == FormatCell(pStrPatten);
                return value;
            }

        }

        public static void MergeCellsFillBackGroundColorList(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn, string value)
        {
            using (var range = worksheet.Cells[startRow, startColumn, endRow, endColumn])
            {
                if (range.Merge != true)
                    range.Merge = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(85, 107, 47));
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Font.SetFromFont(new Font("Arial", 12));
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                range.AutoFitColumns();
                range.Value = value;
            }
        }

        public static void CreateDefaultNamedStyleInWorkBook(ref ExcelPackage ep, string epType) //預先對指定的Excel全Sheet建立Style
        {
            //Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...?
            var namedStyleTitleRow = ep.Workbook.Styles.CreateNamedStyle("Title Row");
            namedStyleTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            namedStyleTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            namedStyleTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyleTitleRow.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(107, 142, 35));
            namedStyleTitleRow.Style.Font.Color.SetColor(Color.White);
            //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            if (epType == "TestFlowProfile")
            {
                var namedStyleTitleRowTf = ep.Workbook.Styles.CreateNamedStyle("TF Title Row");
                namedStyleTitleRowTf.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleTitleRowTf.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleTitleRowTf.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleTitleRowTf.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 176, 240));
                namedStyleTitleRowTf.Style.Font.Color.SetColor(Color.White);
                var namedStyleSubTitleRow = ep.Workbook.Styles.CreateNamedStyle("Sub Title Row");
                namedStyleSubTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleSubTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleSubTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleSubTitleRow.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                namedStyleSubTitleRow.Style.Font.Color.SetColor(Color.White);
                //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var namedStyleOddRow = ep.Workbook.Styles.CreateNamedStyle("Odd Row");
                namedStyleOddRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleOddRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleOddRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleOddRow.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                var namedStyleEvenRow = ep.Workbook.Styles.CreateNamedStyle("Even Row");
                namedStyleEvenRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleEvenRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleEvenRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleEvenRow.Style.Fill.BackgroundColor.SetColor(Color.FloralWhite);

                var namedStyleTestSettingHeader = ep.Workbook.Styles.CreateNamedStyle("Test Setting Header");
                namedStyleTestSettingHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleTestSettingHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleTestSettingHeader.Style.TextRotation = 180;

                var namedStyleHighlightCell = ep.Workbook.Styles.CreateNamedStyle("Highlight");
                namedStyleHighlightCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleHighlightCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleHighlightCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleHighlightCell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            }
        }
    }

    public static class EpplusExtensions
    {
        public static void AddMarcoFromBas(this ExcelPackage excel, string filePath, string moduleName)
        {
            if (File.Exists(filePath))
            {
                ExcelVBAModule module = IsExistModule(excel, moduleName) ? excel.Workbook.VbaProject.Modules[moduleName] : excel.Workbook.VbaProject.Modules.AddModule(moduleName);
                var sb = new StringBuilder();
                using (StreamReader reader = new StreamReader(filePath))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine() + "\r\n";
                        if (!(line.StartsWith("Attribute"))) sb.Append(line);
                    }
                }
                module.Code = sb.ToString();
            }
        }

        private static bool IsExistModule(ExcelPackage excel, string moduleName)
        {
            bool flag = false;
            foreach (var item in excel.Workbook.VbaProject.Modules)
                if (item.Name == moduleName) flag = true;
            return flag;
        }

        #region workbook
        public static string GetMergeCellValue(this ExcelWorksheet worksheet, int rowNum, int colNum)
        {
            string mergedCell = worksheet.MergedCells[rowNum, colNum];
            if (mergedCell == null)
            {
                return worksheet.Cells[rowNum, colNum].Text ?? string.Empty;
            }

            string value = worksheet
                .Cells[new ExcelAddress(mergedCell).Start.Row, new ExcelAddress(mergedCell).Start.Column].Text;
            return value ?? string.Empty;
        }

        public static void CopyWorkSheets(this ExcelWorkbook workbook, List<string> files)
        {
            if (files == null) return;

            ExcelTextFormat format = new ExcelTextFormat
            {
                Delimiter = ',',
                Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                {
                    DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                }
            };
            format.Encoding = new UTF8Encoding();

            foreach (string file in files)
            {
                if (file == null)
                {
                    continue;
                }

                if (Path.GetExtension(file).Equals(".csv", StringComparison.CurrentCultureIgnoreCase))
                {
                    FileInfo fileInfo = new FileInfo(file);
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet =
                            excelPackage.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(file));
                        worksheet.Cells["A1"].LoadFromText(fileInfo, format);
                        workbook.AddSheet(worksheet);
                    }
                }
                else
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            workbook.AddSheet(worksheet);
                        }
                    }
                }
            }
        }

        public static void AddSheet(this ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == worksheet.Name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                workbook.Worksheets[worksheet.Name].Cells.Clear();
            }
            else
            {
                workbook.Worksheets.Add(worksheet.Name, worksheet);
            }

            workbook.Worksheets.MoveBefore(worksheet.Name, workbook.Worksheets[1].Name);
        }

        public static void DeleteSheet(this ExcelWorkbook workbook, string name)
        {
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (name.Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    workbook.Worksheets.Delete(sheet);
                    break;
                }
            }
        }

        public static ExcelWorksheet AddSheet(this ExcelWorkbook workbook, string name)
        {
            bool isExist = false;
            foreach (ExcelWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == name)
                {
                    isExist = true;
                }
            }

            if (isExist)
            {
                workbook.Worksheets[name].Cells.Clear();
            }
            else
            {
                workbook.Worksheets.Add(name);
            }

            workbook.Worksheets.MoveBefore(name, workbook.Worksheets[1].Name);
            return workbook.Worksheets[1];
        }

        private static void AddHeaderStyle(this ExcelWorkbook workbook)
        {
            foreach (ExcelNamedStyleXml item in workbook.Styles.NamedStyles)
            {
                if (item.Name == "Header")
                {
                    return;
                }
            }

            ExcelNamedStyleXml namedStyle = workbook.Styles.CreateNamedStyle("Header");
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);
        }
        #endregion

        #region worksheet
        public static DataTable ReadSheetToDataSet(this ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return null;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            DataTable dt = new DataTable(worksheet.Name);
            if (rowCount > 0)
            {
                object objCellValue;
                object cellValue;
                for (int j = 0; j < columnCount; j++) //設定DataTable列名
                {
                    objCellValue = worksheet.Cells[1, j + 1].Value;
                    cellValue = objCellValue == null ? "" : objCellValue.ToString();
                    dt.Columns.Add(cellValue.ToString(), typeof(object));
                }
                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        objCellValue = worksheet.Cells[i, j].Value;
                        cellValue = objCellValue ?? "";
                        dr[j - 1] = cellValue;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        public static void MergeColumn(this ExcelWorksheet worksheet, int colNum, bool isHeaders = true)
        {
            object[,] array = (object[,])worksheet.Cells[1, colNum, worksheet.Dimension.End.Row, colNum].Value;
            if (array.Length == 1)
            {
                return;
            }

            int lastIndex = 0;
            int start = isHeaders ? 1 : 0;
            for (int i = start; i < array.Length; i++)
            {
                for (int j = i + 1; j < array.Length; j++)
                {
                    if (!array[i, 0].ToString().Equals(array[j, 0].ToString(), StringComparison.CurrentCulture))
                    {
                        lastIndex = j - 1;
                        break;
                    }

                    if (j == array.Length - 1)
                    {
                        lastIndex = j;
                        break;
                    }
                }

                if (lastIndex - i > 0)
                {
                    worksheet.Cells[i + 1, colNum, lastIndex + 1, colNum].Merge = true;
                    worksheet.Cells[i + 1, colNum, lastIndex + 1, colNum].Style.VerticalAlignment =
                        ExcelVerticalAlignment.Center;
                    i = lastIndex;
                }
            }
        }

        public static void SetFormula(this ExcelWorksheet worksheet, int colNum, bool isHeaders = true)
        {
            int start = isHeaders ? 2 : 1;
            for (int i = start; i <= worksheet.Dimension.End.Row; i++)
            {
                if (worksheet.Cells[i, colNum].Value != null)
                    SetHyperLinkFormat(worksheet.Cells[i, colNum]);
            }
        }

        public static void SetHyperLinkFormat(this ExcelRange excelRange)
        {
            excelRange.Formula = excelRange.Value.ToString();
            excelRange.Style.Font.UnderLine = true;
            excelRange.Style.Font.Color.SetColor(Color.Blue);
        }

        public static void SetHeaderStyle(this ExcelWorksheet worksheet)
        {
            worksheet.Workbook.AddHeaderStyle();
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.Start.Row, worksheet.Dimension.End.Column].StyleName =
                "Header";
        }

        public static void SetHairBorder(this ExcelWorksheet worksheet)
        {
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Top.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Bottom.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Left.Style =
                ExcelBorderStyle.Hair;
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                    worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Border.Right.Style =
                ExcelBorderStyle.Hair;
        }

        public static void ExportSheet(this ExcelWorksheet worksheet, string txtFileName, string delimiter = "\t")
        {
            var maxRow = worksheet.Dimension.Rows;
            var maxCol = worksheet.Dimension.Columns;
            StreamWriter sw = File.CreateText(txtFileName);
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                List<string> row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    var value = worksheet.Cells[iRow, iCol].Value ?? "";
                    row.Add(value.ToString());
                }
                sw.WriteLine(string.Join(delimiter, row));
            }
            sw.Close();
        }
        #endregion

        #region Print
        public static int PrintExcelRow<T>(this ExcelRange range, T[] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            List<object[]> arrays = new List<object[]>();
            object[] array = data.OfType<object>().ToArray();
            arrays.Add(array);
            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
            return startRow + arrays.Count();
        }

        public static void PrintExcelCol<T>(this ExcelRange range, T[] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            T[,] dataArray = new T[data.GetLength(0), 1];
            for (int i = 0; i < dataArray.GetLength(0); i++)
            {
                for (int j = 0; j < dataArray.GetLength(1); j++)
                {
                    dataArray[i, j] = data[i];
                }
            }

            worksheet.Cells[startRow, startCol].LoadFromCollection(data);
        }

        public static int PrintExcelRange<T>(this ExcelRange range, T[,] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            List<object[]> arrays = new List<object[]>();
            for (int i = 0; i < data.GetLength(0); i++)
            {
                object[] array = new object[data.GetLength(1)];
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    array[j] = data[i, j];
                }
                arrays.Add(array);
            }

            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
            return startRow + data.GetLength(0);
        }

        public static void PrintExcelColByList<T>(this ExcelRange range, List<List<T>> list)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            int rowCnt = list[0].Count;
            for (int index = startCol; index < list.Count; index++)
            {
                List<T> item = list[index];
                worksheet.Cells[startRow, index, startRow + rowCnt - 1, index].LoadFromCollection(item);
            }
        }

        public static int PrintExcelRowByList<T>(this ExcelRange range, List<List<T>> list)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            if (list.Count == 0)
            {
                return 0;
            }

            int rowCnt = list.Count;
            int colCnt = list[0].Count;

            DataTable dataTable = new DataTable();
            for (int i = 0; i < colCnt; i++)
            {
                dataTable.Columns.Add();
            }

            for (int i = 0; i < rowCnt; i++)
            {
                DataRow row = dataTable.NewRow();
                colCnt = list[i].Count;
                for (int j = 0; j < colCnt; j++)
                {
                    row[j] = list[i][j];
                }

                dataTable.Rows.Add(row);
            }

            worksheet.Cells[startRow, startCol, startRow + dataTable.Rows.Count - 1,
                startCol + dataTable.Columns.Count - 1].LoadFromDataTable(dataTable, false);
            return startRow + rowCnt;
        }

        public static void AddHyperLink(this ExcelRange range, string sheetName2, int x2, int y2)
        {
            if (x2 > 0 && y2 > 0)
            {
                string cellPosBase = ExcelCellBase.GetAddress(x2, y2);
                range.Hyperlink = new ExcelHyperLink(sheetName2 + "!" + cellPosBase, range.Text);
                range.StyleName = "HyperLink";
                range.Style.Font.UnderLine = true;
            }
        }
        #endregion
    }
}
