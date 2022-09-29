using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace CommonLib.Extension
{
    public static class EpplusExtensions
    {
        #region workbook

        public static string GetMergeCellValue(this ExcelWorksheet worksheet, int rowNum, int colNum)
        {
            var mergedCell = worksheet.MergedCells[rowNum, colNum];
            if (mergedCell == null) return worksheet.Cells[rowNum, colNum].Text ?? string.Empty;

            var value = worksheet
                .Cells[new ExcelAddress(mergedCell).Start.Row, new ExcelAddress(mergedCell).Start.Column].Text;
            return value ?? string.Empty;
        }

        public static void CopyWorkSheets(this ExcelWorkbook workbook, List<string> files)
        {
            if (files == null) return;

            var format = new ExcelTextFormat
            {
                Delimiter = ',',
                Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString())
                {
                    DateTimeFormat = { ShortDatePattern = "dd-mm-yyyy" }
                }
            };
            format.Encoding = new UTF8Encoding();

            foreach (var file in files)
            {
                if (file == null) continue;

                if (Path.GetExtension(file).Equals(".csv", StringComparison.CurrentCultureIgnoreCase))
                {
                    var fileInfo = new FileInfo(file);
                    using (var excelPackage = new ExcelPackage())
                    {
                        var worksheet =
                            excelPackage.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(file));
                        worksheet.Cells["A1"].LoadFromText(fileInfo, format);
                        workbook.AddSheet(worksheet);
                    }
                }
                else
                {
                    using (var package = new ExcelPackage(new FileInfo(file)))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets) workbook.AddSheet(worksheet);
                    }
                }
            }
        }

        public static void AddSheet(this ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            var isExist = false;
            foreach (var sheet in workbook.Worksheets)
                if (sheet.Name == worksheet.Name)
                    isExist = true;

            if (isExist)
                workbook.Worksheets[worksheet.Name].Cells.Clear();
            else
                workbook.Worksheets.Add(worksheet.Name, worksheet);

            workbook.Worksheets.MoveBefore(worksheet.Name, workbook.Worksheets[1].Name);
        }

        public static void DeleteSheet(this ExcelWorkbook workbook, string name)
        {
            foreach (var sheet in workbook.Worksheets)
                if (name.Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    workbook.Worksheets.Delete(sheet);
                    break;
                }
        }

        public static ExcelWorksheet AddSheet(this ExcelWorkbook workbook, string name)
        {
            var isExist = false;
            foreach (var sheet in workbook.Worksheets)
                if (sheet.Name == name)
                    isExist = true;

            if (isExist)
                workbook.Worksheets[name].Cells.Clear();
            else
                workbook.Worksheets.Add(name);

            workbook.Worksheets.MoveBefore(name, workbook.Worksheets[1].Name);
            return workbook.Worksheets[1];
        }

        private static void AddHeaderStyle(this ExcelWorkbook workbook)
        {
            foreach (var item in workbook.Styles.NamedStyles)
                if (item.Name == "Header")
                    return;

            var namedStyle = workbook.Styles.CreateNamedStyle("Header");
            namedStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyle.Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);
        }

        #endregion

        #region worksheet
        public static string GetMergedCellValue(this ExcelWorksheet sheet, int rowNumber, int columnNumber)
        {
            var range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null
                ? GetCellValue(sheet, rowNumber, columnNumber)
                : GetCellValue(sheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column);
        }

        public static string GetCellValue(this ExcelWorksheet sheet, int row, int column)
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


        public static Dictionary<string, int> GetHeaderOrder(this ExcelWorksheet sheet, int startRow = 1)
        {
            var headerOrder = new Dictionary<string, int>();
            if (sheet.Dimension == null)
                return headerOrder;
            var endCol = sheet.Dimension.End.Column;
            for (var i = 1; i <= endCol; i++)
                if (sheet.Cells[startRow, i].Value != null)
                {
                    var header = sheet.Cells[startRow, i].Value.ToString().Trim();
                    if (!headerOrder.ContainsKey(header))
                        headerOrder.Add(header, i);
                }

            return headerOrder;
        }



        public static string GetMergedCellValueAndAddress(this ExcelWorksheet sheet, int rowNumber, int columnNumber,
            ref string address)
        {
            var range = sheet.MergedCells[rowNumber, columnNumber];
            return range == null
                ? GetCellValueAndAddress(sheet, rowNumber, columnNumber, ref address)
                : GetCellValueAndAddress(sheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column,
                    ref address);
        }

        public static string GetCellValueAndAddress(this ExcelWorksheet sheet, int row, int column, ref string address)
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
                var range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Value.ToString();
            }

            if (sheet.Cells[row, column].Text != null)
            {
                var range = sheet.Cells[row, column];
                address = range.Address;
                return sheet.Cells[row, column].Text;
            }

            return "";
        }

        public static DataTable ExportToDataSet(this ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return null;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            var dt = new DataTable(worksheet.Name);
            if (rowCount > 0)
            {
                object objCellValue;
                object cellValue;
                for (var j = 0; j < columnCount; j++)
                {
                    objCellValue = worksheet.Cells[1, j + 1].Value;
                    cellValue = objCellValue == null ? "" : objCellValue.ToString();
                    dt.Columns.Add(cellValue.ToString(), typeof(object));
                }

                for (var i = 2; i <= rowCount; i++)
                {
                    var dr = dt.NewRow();
                    for (var j = 1; j <= columnCount; j++)
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

        public static void ExportToTxt(this ExcelWorksheet worksheet, string filePath, string delimiter = "\t",
            string[,] pColumnNameFixedValue = null,
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
                using (var sw = new StreamWriter(filePath))
                {
                    for (var i = 1; i <= rowCount; i++)
                    {
                        var arr = new List<string>();
                        for (var j = 1; j <= columnCount; j++)
                        {
                            //skip specified column, can't export specified column
                            if (skipColumnList != null && skipColumnList.Contains(j))
                                continue;
                            var objCellValue = worksheet.Cells[i, j].Value;
                            var cellValue = objCellValue ?? "";
                            arr.Add(cellValue.ToString());
                        }

                        if (pColumnNameFixedValue != null)
                        {
                            //it's the last column of the first row
                            if (i == 1)
                                for (var lIntRow = 0;
                                     lIntRow < pColumnNameFixedValue.GetLength(0);
                                     lIntRow++) //getlength() 1->0, by Ze
                                {
                                    var lIntIndex = int.Parse(pColumnNameFixedValue[lIntRow, 2]);
                                    arr.Insert(lIntIndex, pColumnNameFixedValue[lIntRow, 0]);
                                }
                            else
                                //it's the last column of non-first row
                                for (var lIntColumn = 0;
                                     lIntColumn < pColumnNameFixedValue.GetLength(0);
                                     lIntColumn++) //getlength() 1->0, by Ze
                                {
                                    var lIntIndex = int.Parse(pColumnNameFixedValue[lIntColumn, 2]);
                                    arr.Insert(lIntIndex, pColumnNameFixedValue[lIntColumn, 1]);
                                }
                        }

                        sw.WriteLine(string.Join(delimiter, arr));
                    }

                    if (extraRows != null)
                    {
                        var extraRowCount = extraRows.Values.Select(o => o.Count).Max();
                        for (var i = 1; i <= extraRowCount; i++)
                        {
                            var extraArr = new string[extraRows.Keys.Max() + 1].ToList();
                            foreach (var index in extraRows.Keys)
                                if (extraRows[index].Count >= i)
                                    extraArr[index] = extraRows[index][i - 1];
                            sw.WriteLine(string.Join(delimiter, extraArr));
                        }
                    }
                }
        }

        public static void VddLevelExportToTxt(this ExcelWorksheet worksheet, string filePath, int seqIndex,
            string delimiter = "\t", string[,] pColumnNameFixedValue = null)
        {
            if (worksheet == null) return;
            if (worksheet.Dimension == null) return;
            var columnCount = worksheet.Dimension.End.Column;
            var rowCount = worksheet.Dimension.End.Row;
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            if (rowCount > 0)
                using (var sw = new StreamWriter(filePath))
                {
                    for (var i = 1; i <= rowCount; i++)
                    {
                        var seqColumnValue = worksheet.Cells[i, seqIndex].Value;
                        var seqColumnStr = seqColumnValue ?? "";
                        if (seqColumnStr.ToString().Equals("x", StringComparison.CurrentCultureIgnoreCase))
                            continue;
                        var arr = new List<string>();
                        for (var j = 1; j <= columnCount; j++)
                        {
                            var objCellValue = worksheet.Cells[i, j].Value;
                            var cellValue = objCellValue ?? "";
                            arr.Add(cellValue.ToString());
                        }

                        //if (pColumnNameFixedValue != null)
                        //{
                        //    //it's the last column of the first row
                        //    if (i == 1)
                        //        for (var lIntRow = 0;
                        //             lIntRow < pColumnNameFixedValue.GetLength(0);
                        //             lIntRow++) //getlength() 1->0, by Ze
                        //        {
                        //            var lIntIndex = int.Parse(pColumnNameFixedValue[lIntRow, 2]);
                        //            arr.Insert(lIntIndex, pColumnNameFixedValue[lIntRow, 0]);
                        //        }
                        //    else
                        //        //it's the last column of non-first row
                        //        for (var lIntColumn = 0;
                        //             lIntColumn < pColumnNameFixedValue.GetLength(0);
                        //             lIntColumn++) //getlength() 1->0, by Ze
                        //        {
                        //            var lIntIndex = int.Parse(pColumnNameFixedValue[lIntColumn, 2]);
                        //            //if column 'WS Bump Name' in VDD_Levels is empty, value is assigned to empty too.
                        //            var objCellValue = worksheet.Cells[i, 1].Value;
                        //            var cellValue = objCellValue ?? "";
                        //            if (string.IsNullOrEmpty(cellValue.ToString()))
                        //            {
                        //                arr.Insert(lIntIndex, "");
                        //            }
                        //            else
                        //            {
                        //                //The Power level sheet if pin = DC30 then BW need set to 500.
                        //                if (objCellValue.ToString().Trim().EndsWith("_DC30",
                        //                        StringComparison.CurrentCultureIgnoreCase)
                        //                    && (pColumnNameFixedValue[lIntColumn, 0] == "BW_LowCap"
                        //                        || pColumnNameFixedValue[lIntColumn, 0] == "BW_HighCap"
                        //                        || pColumnNameFixedValue[lIntColumn, 0] == "CPBorad_BW_LowCap"
                        //                        || pColumnNameFixedValue[lIntColumn, 0] == "CPBorad_BW_HighCap"))
                        //                    arr.Insert(lIntIndex, "500");
                        //                else
                        //                    arr.Insert(lIntIndex, pColumnNameFixedValue[lIntColumn, 1]);
                        //            }
                        //        }
                        //}

                        sw.WriteLine(string.Join(delimiter, arr));
                    }
                }
        }

        public static void MergeColumn(this ExcelWorksheet worksheet, int colNum, bool isHeaders = true)
        {
            var array = (object[,])worksheet.Cells[1, colNum, worksheet.Dimension.End.Row, colNum].Value;
            if (array.Length == 1) return;

            var lastIndex = 0;
            var start = isHeaders ? 1 : 0;
            for (var i = start; i < array.Length; i++)
            {
                for (var j = i + 1; j < array.Length; j++)
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
            var start = isHeaders ? 2 : 1;
            for (var i = start; i <= worksheet.Dimension.End.Row; i++)
                if (worksheet.Cells[i, colNum].Value != null)
                    SetHyperLinkFormat(worksheet.Cells[i, colNum]);
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
            var worksheet = range.Worksheet;
            var startRow = range.Start.Row;
            var startCol = range.Start.Column;
            var arrays = new List<object[]>();
            var array = data.OfType<object>().ToArray();
            arrays.Add(array);
            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
        }

        public static void PrintExcelCol<T>(this ExcelRange range, T[] data)
        {
            var worksheet = range.Worksheet;
            var startRow = range.Start.Row;
            var startCol = range.Start.Column;
            var dataArray = new T[data.GetLength(0), 1];
            for (var i = 0; i < dataArray.GetLength(0); i++)
                for (var j = 0; j < dataArray.GetLength(1); j++)
                    dataArray[i, j] = data[i];

            worksheet.Cells[startRow, startCol].LoadFromCollection(data);
        }

        public static void PrintExcelRange<T>(this ExcelRange range, T[,] data)
        {
            var worksheet = range.Worksheet;
            var startRow = range.Start.Row;
            var startCol = range.Start.Column;
            var arrays = new List<object[]>();
            for (var i = 0; i < data.GetLength(0); i++)
            {
                var array = new object[data.GetLength(1)];
                for (var j = 0; j < data.GetLength(1); j++) array[j] = data[i, j];

                arrays.Add(array);
            }

            worksheet.Cells[startRow, startCol].LoadFromArrays(arrays);
        }

        public static void PrintExcelColByList<T>(this ExcelRange range, List<List<T>> list)
        {
            var worksheet = range.Worksheet;
            var startRow = range.Start.Row;
            var startCol = range.Start.Column;
            var rowCnt = list[0].Count;
            for (var index = startCol; index < list.Count; index++)
            {
                var item = list[index];
                worksheet.Cells[startRow, index, startRow + rowCnt - 1, index].LoadFromCollection(item);
            }
        }

        public static int PrintExcelRowByList<T>(this ExcelRange range, List<List<T>> list)
        {
            var worksheet = range.Worksheet;
            var startRow = range.Start.Row;
            var startCol = range.Start.Column;
            if (list.Count == 0) return 0;

            var rowCnt = list.Count;
            var colCnt = list[0].Count;

            var dataTable = new DataTable();
            for (var i = 0; i < colCnt; i++) dataTable.Columns.Add();

            for (var i = 0; i < rowCnt; i++)
            {
                var row = dataTable.NewRow();
                colCnt = list[i].Count;
                for (var j = 0; j < colCnt; j++) row[j] = list[i][j];

                dataTable.Rows.Add(row);
            }

            worksheet.Cells[startRow, startCol, startRow + dataTable.Rows.Count - 1,
                startCol + dataTable.Columns.Count - 1].LoadFromDataTable(dataTable, false);
            return startRow + rowCnt;
        }

        #endregion
    }
}