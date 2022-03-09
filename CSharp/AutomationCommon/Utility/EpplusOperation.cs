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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace AutomationCommon.Utility
{
    public class EpplusOperation
    {
        #region Operation on excel
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
        public static string CopyExcelWorkBook(string sourcePath, string targetPath, string excelExtension = "", bool clearStyle = false)
        {
            if (File.Exists(targetPath))
                File.Delete(targetPath);
            File.Copy(sourcePath, targetPath);
            return targetPath;
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

        public static string GetMergedCellValueAndAddress(ExcelWorksheet sheet, int rowNumber, int columnNumber, ref string address)
        {
            string range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null ?
                GetCellValueAndAddress(sheet, rowNumber, columnNumber, ref address) :
                GetCellValueAndAddress(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column), ref address);
        }

        public static string GetCellValue(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet == null)
                return "";
            if (row <= 0 || column <= 0) return "";
            // if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            //return sheet.Cells[row, column].Formula;
            if (sheet.Cells[row, column] == null)
                return "";
            if (sheet.Cells[row, column].Value != null)
                return sheet.Cells[row, column].Value.ToString();
            if (sheet.Cells[row, column].Text != null)
                return sheet.Cells[row, column].Text;
            return "";
        }

        public static string GetCellValueAndAddress(ExcelWorksheet sheet, int row, int column, ref string address)
        {
            address = string.Empty;
            if (sheet == null)
                return "";
            if (row <= 0 || column <= 0) return "";
            // if (!string.IsNullOrEmpty(sheet.Cells[row, column].Formula))
            //return sheet.Cells[row, column].Formula;
            if (sheet.Cells[row, column] == null)
                return "";
            if (sheet.Cells[row, column].Value != null)
            {
                ExcelRange range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Value.ToString();
            }
            if (sheet.Cells[row, column].Text != null)
            {
                ExcelRange range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Text;
            }
            return "";
        }

        public static string GetCellValueOld(ExcelWorksheet wSheet, int row, int column)
        {
            if (row <= 0 || column <= 0) return "";
            if (wSheet.Cells[row, column].Value != null)
                return wSheet.Cells[row, column].Value.ToString().Trim();
            return "";
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
        public static DataTable ExportToDataSet(this ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return null;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            DataTable dt = new DataTable(worksheet.Name);
            if (rowCount > 0)
            {
                object objCellValue;
                object cellValue;
                for (int j = 0; j < columnCount; j++)
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

        public static void ExportToTxt(this ExcelWorksheet worksheet, string filePath, string delimiter = "\t", string[,] p_ColumnNameFixedValue = null, 
            List<int> skipColumnList = null, Dictionary<int, List<string>> extraRows = null)
        {
            if (worksheet == null) return;
            if (worksheet.Dimension == null) return;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            if (rowCount > 0)
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        List<string> arr = new List<string>();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            //skip specified column, can't export specified column
                            if (skipColumnList != null && skipColumnList.Contains(j))
                                continue;
                            var objCellValue = worksheet.Cells[i, j].Value;
                            var cellValue = objCellValue ?? "";
                            arr.Add(cellValue.ToString());
                        }

                        if (p_ColumnNameFixedValue != null)
                        {
                            //it's the last column of the first row
                            if (i == 1)
                            {
                                for (int l_intRow = 0; l_intRow < p_ColumnNameFixedValue.GetLength(0); l_intRow++)//getlength() 1->0, by Ze
                                {
                                    int l_intIndex = int.Parse(p_ColumnNameFixedValue[l_intRow, 2]);
                                    arr.Insert(l_intIndex, p_ColumnNameFixedValue[l_intRow, 0]);
                                }
                            }
                            else
                            {
                                //it's the last column of non-first row
                                for (int l_intColumn = 0; l_intColumn < p_ColumnNameFixedValue.GetLength(0); l_intColumn++)//getlength() 1->0, by Ze
                                {
                                    int l_intIndex = int.Parse(p_ColumnNameFixedValue[l_intColumn, 2]);
                                    arr.Insert(l_intIndex, p_ColumnNameFixedValue[l_intColumn, 1]);
                                }
                            }
                        }
                        else
                        {
                            //do nothing
                        }
                        sw.WriteLine(string.Join(delimiter, arr));
                    }
                    if (extraRows != null)
                    {
                        int extraRowCount = extraRows.Values.Select(o => o.Count).Max();
                        for (int i = 1; i <= extraRowCount; i++)
                        {
                            List<string> extraArr = (new string[extraRows.Keys.Max() + 1]).ToList();
                            foreach (var index in extraRows.Keys)
                            {
                                if (extraRows[index].Count >= i)
                                {
                                    extraArr[index] = extraRows[index][i - 1];
                                }
                            }
                            sw.WriteLine(string.Join(delimiter, extraArr));
                        }
                    }
                }
            }
        }

        public static void VddLevelExportToTxt(this ExcelWorksheet worksheet, string filePath, int seqIndex, string delimiter = "\t", string[,] p_ColumnNameFixedValue = null, List<int> skipColumnList = null)
        {
            if (worksheet == null) return;
            if (worksheet.Dimension == null) return;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            if (rowCount > 0)
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        var seqColumnValue = worksheet.Cells[i, seqIndex].Value;
                        var seqColumnStr = seqColumnValue ?? "";
                        if (seqColumnStr.ToString().Equals("x", StringComparison.CurrentCultureIgnoreCase))
                            continue;
                        List<string> arr = new List<string>();
                        for (int j = 1; j <= columnCount; j++)
                        {
                            var objCellValue = worksheet.Cells[i, j].Value;
                            var cellValue = objCellValue ?? "";
                            arr.Add(cellValue.ToString());
                        }

                        if (p_ColumnNameFixedValue != null)
                        {
                            //it's the last column of the first row
                            if (i == 1)
                            {
                                for (int l_intRow = 0; l_intRow < p_ColumnNameFixedValue.GetLength(0); l_intRow++)//getlength() 1->0, by Ze
                                {
                                    int l_intIndex = int.Parse(p_ColumnNameFixedValue[l_intRow, 2]);
                                    arr.Insert(l_intIndex, p_ColumnNameFixedValue[l_intRow, 0]);
                                }
                            }
                            else
                            {
                                //it's the last column of non-first row
                                for (int l_intColumn = 0; l_intColumn < p_ColumnNameFixedValue.GetLength(0); l_intColumn++)//getlength() 1->0, by Ze
                                {
                                    int l_intIndex = int.Parse(p_ColumnNameFixedValue[l_intColumn, 2]);
                                    //if column 'WS Bump Name' in VDD_Levels is empty, value is assigned to empty too.
                                    var objCellValue = worksheet.Cells[i, 1].Value;
                                    var cellValue = objCellValue ?? "";
                                    if (string.IsNullOrEmpty(cellValue.ToString()))
                                    {
                                        arr.Insert(l_intIndex, "");
                                    }
                                    else
                                    {
                                        //The Power level sheet if pin = DC30 then BW need set to 500.
                                        if (objCellValue.ToString().Trim().EndsWith("_DC30", StringComparison.CurrentCultureIgnoreCase)
                                            && (p_ColumnNameFixedValue[l_intColumn, 0] == "BW_LowCap"
                                            || p_ColumnNameFixedValue[l_intColumn, 0] == "BW_HighCap"
                                            || p_ColumnNameFixedValue[l_intColumn, 0] == "CPBorad_BW_LowCap"
                                            || p_ColumnNameFixedValue[l_intColumn, 0] == "CPBorad_BW_HighCap"))
                                        {
                                            arr.Insert(l_intIndex, "500");
                                        }
                                        else
                                        {
                                            arr.Insert(l_intIndex, p_ColumnNameFixedValue[l_intColumn, 1]);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            //do nothing
                        }
                        if (skipColumnList != null && skipColumnList.Count > 0)
                        {
                            List<string> newArr = new List<string>();
                            for (int j = 0; j < arr.Count; j++)
                            {
                                if (skipColumnList.Contains(j))
                                    continue;
                                newArr.Add(arr[j]);
                            }
                            sw.WriteLine(string.Join(delimiter, newArr));
                        }
                        else
                        {
                            sw.WriteLine(string.Join(delimiter, arr));
                        }
                    }
                }
            }
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