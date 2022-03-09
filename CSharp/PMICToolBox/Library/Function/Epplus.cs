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

namespace Library.Function
{
    public static class Epplus
    {
        #region Basic

        public static bool IsExcelOpened(string file)
        {
            if (!File.Exists(file))
            {
                return false;
            }

            try
            {
                Stream s = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        public static void KillExcel()
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
        }

        #endregion

        #region Data convertor

        public static string[,] WorkSheetToStringArray(ExcelWorksheet worksheet)
        {
            if (worksheet?.Dimension == null)
            {
                return null;
            }

            string[,] array = new string[worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    array[row, col] = worksheet.GetMergeCellValue(row, col);
                }
            }

            return array;
        }

        #endregion

        #region Operation

        public static string ConvertUnit(string value)
        {
            if (Regex.IsMatch(value.Trim(' '), @"^[-+]?\d+(\.\d+)?\D+$"))
            {
                return ParsePrefix(value);
            }

            return value;
        }

        #endregion

        public static string ParsePrefix(string value)
        {
            string[] superSuffix = { "K", "M", "G", "T", "P", "A" };
            string[] subSuffix = { "m", "u", "n", "p", "f", "a" };

            value = value.Replace("*", "");
            string unit = Regex.Replace(value, @"[-+]?\d+(\.\d+)?", "").Trim();
            if (unit.Length > 0)
            {
                char c = unit[0];
                char suffixChar;
                foreach (string s in subSuffix)
                {
                    if (c.ToString() == s)
                    {
                        suffixChar = c;
                        string num = value.Substring(0, value.IndexOf(suffixChar));
                        return (Convert.ToDouble(num) / Math.Pow(1000,
                                    subSuffix.ToList().IndexOf(suffixChar.ToString()) + 1)).ToString("G");
                    }
                }

                foreach (string s in superSuffix)
                {
                    if (c.ToString().ToLower() == s.ToLower())
                    {
                        suffixChar = s[0];
                        string num = value.Substring(0, value.IndexOf(c));
                        double multi = Math.Pow(1000, superSuffix.ToList().IndexOf
                                                          (suffixChar.ToString()) + 1);
                        return (Convert.ToDouble(num) * multi).ToString("G");
                    }
                }
            }

            return value;
        }

        public static string AddPrefix(double value, string unit = "")
        {
            string[] superSuffix = { "K", "M", "G", "T", "P", "A" };
            string[] subSuffix = { "m", "u", "n", "p", "f", "a" };
            double v = value;
            int exp = 0;
            while (v - Math.Floor(v) > 0)
            {
                if (exp >= 18)
                {
                    break;
                }

                exp += 3;
                v *= 1000;
                v = Math.Round(v, 12);
            }

            while (Math.Floor(v).ToString(CultureInfo.InvariantCulture).Length > 3)
            {
                if (exp <= -18)
                {
                    break;
                }

                exp -= 3;
                v /= 1000;
                v = Math.Round(v, 12);
            }

            if (exp > 0)
            {
                return v + subSuffix[exp / 3 - 1] + unit;
            }

            if (exp < 0)
            {
                return v + superSuffix[-exp / 3 - 1] + unit;
            }

            return v + unit;
        }

        #region VBT module

        public static void AddMarcoFromBas(ExcelPackage excel, string file, string moduleName)
        {
            if (File.Exists(file))
            {
                ExcelVBAModule module = IsExistModule(excel, moduleName)
                    ? excel.Workbook.VbaProject.Modules[moduleName]
                    : excel.Workbook.VbaProject.Modules.AddModule(moduleName);
                StringBuilder sb = new StringBuilder();
                using (StreamReader reader = new StreamReader(file))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine() + "\r\n";
                        if (!line.StartsWith("Attribute"))
                        {
                            sb.Append(line);
                        }
                    }
                }

                module.Code = sb.ToString();
            }
        }

        private static bool IsExistModule(ExcelPackage excel, string moduleName)
        {
            bool flag = false;
            foreach (ExcelVBAModule item in excel.Workbook.VbaProject.Modules)
            {
                if (item.Name == moduleName)
                {
                    flag = true;
                }
            }

            return flag;
        }

        #endregion
    }

    public static class EpplusExtensions
    {
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
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(file));
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

            if (workbook.Worksheets.Count > 1)
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

            if (workbook.Worksheets.Count > 1)
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
                    if (array[i, 0] != array[j, 0])
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
                {
                    worksheet.Cells[i, colNum].Formula = worksheet.Cells[i, colNum].Value.ToString();
                    worksheet.Cells[i, colNum].Style.Font.UnderLine = true;
                    worksheet.Cells[i, colNum].Style.Font.Color.SetColor(Color.Blue);
                }
            }
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

        #endregion

        #region Print

        public static void PrintExcelRow<T>(this ExcelRange range, T[] data)
        {
            ExcelWorksheet worksheet = range.Worksheet;
            int startRow = range.Start.Row;
            int startCol = range.Start.Column;
            List<object[]> arrays = new List<object[]>();
            object[] array = data.OfType<object>().ToArray();
            arrays.Add(array);
            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
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

        public static void PrintExcelRange<T>(this ExcelRange range, T[,] data)
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

        #endregion
    }
}