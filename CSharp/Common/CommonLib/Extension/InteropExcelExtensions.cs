using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static class InteropExcelExtensions
    {
        #region workbook

        public static void AddMacro(this Workbook workbook, string basFilePath)
        {
            var basName = Path.GetFileNameWithoutExtension(basFilePath);
            if (!File.Exists(basFilePath)) return;

            if (!workbook.IsVbComponentExist(basName))
            {
                var vBComponent = workbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                var codeModule = vBComponent.CodeModule;
                codeModule.AddFromFile(basFilePath);
                Marshal.ReleaseComObject(codeModule);
            }
        }

        public static bool IsVbComponentExist(this Workbook workbook, string codeModule)
        {
            foreach (VBComponent component in workbook.VBProject.VBComponents)
                if (component.Type == vbext_ComponentType.vbext_ct_StdModule ||
                    component.Type == vbext_ComponentType.vbext_ct_ClassModule)
                    if (codeModule != null &&
                        codeModule.Equals(component.CodeModule.Name, StringComparison.OrdinalIgnoreCase))
                        return true;

            return false;
        }

        public static VBComponent GetVbComponents(this Workbook workbook, string codeModule)
        {
            foreach (VBComponent component in workbook.VBProject.VBComponents)
                if (component.CodeModule.Name == codeModule)
                    //&& component.Type == vbext_ComponentType.vbext_ct_StdModule)
                    return component;
            return workbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }

        public static void ExportVbt(this Workbook workbook, string targetPath)
        {
            var project = workbook.VBProject;
            foreach (VBComponent component in project.VBComponents)
                if (component != null && component.CodeModule.CountOfLines > 0)
                {
                    string file;
                    if (component.Type == vbext_ComponentType.vbext_ct_ClassModule)
                    {
                        file = Path.Combine(targetPath, component.CodeModule.Name + @".cls");
                        component.Export(file);
                    }

                    if (component.Type == vbext_ComponentType.vbext_ct_StdModule)
                    {
                        file = Path.Combine(targetPath, component.CodeModule.Name + @".bas");
                        component.Export(file);
                    }
                }

            Marshal.ReleaseComObject(project);
        }

        public static void ReplaceVbtStr(this Workbook workbook, string oldString, string newString)
        {
            var project = workbook.VBProject;
            foreach (VBComponent component in project.VBComponents)
                if (component.Type == vbext_ComponentType.vbext_ct_StdModule ||
                    component.Type == vbext_ComponentType.vbext_ct_ClassModule)
                {
                    var module = component.CodeModule;
                    if (module.CountOfLines > 0)
                    {
                        var lines = module.Lines[1, module.CountOfLines]
                            .Split(new[] { "\r\n" }, StringSplitOptions.None);
                        for (var i = 0; i < lines.Length; i++)
                            if (Regex.IsMatch(lines[i], oldString, RegexOptions.IgnoreCase))
                            {
                                lines[i] = Regex.Replace(lines[i], oldString, newString, RegexOptions.IgnoreCase);
                                module.ReplaceLine(i + 1, lines[i]);
                            }
                    }
                }

            Marshal.ReleaseComObject(project);
        }

        public static Worksheet GetSheet(this Workbook workbook, string name)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
                if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return worksheet;

            return null;
        }

        public static Worksheet AddSheet(this Workbook workbook, string name)
        {
            if (workbook.IsSheetExist(name))
                workbook.Worksheets[name].Delete();

            Worksheet newSheet =
                workbook.Worksheets.Add(workbook.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            newSheet.Name = name;
            return newSheet;
        }

        public static bool IsSheetExist(this Workbook workbook, string name)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
                if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return true;

            return false;
        }

        public static void SaveAsXlsm(this Workbook workbook)
        {
            var oldFilePath = workbook.FullName;
            var extension = Path.GetExtension(oldFilePath);
            if (extension == null) return;

            if (extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)) return;

            var newFilePath = Path.ChangeExtension(oldFilePath, "xlsm");
            workbook.SaveAs(newFilePath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                null, null, false, false, XlSaveAsAccessMode.xlNoChange,
                false, false);
            if (oldFilePath == null) throw new ArgumentNullException();

            File.Delete(oldFilePath);
        }

        public static void AddHeaderStyle(this Workbook workbook)
        {
            foreach (Style item in workbook.Styles)
                if (item.Name == "Header")
                    return;

            var style = workbook.Styles.Add("Header");
            style.Interior.Color = Color.YellowGreen;
            style.IncludeNumber = true;
            style.IncludeFont = true;
            style.IncludeAlignment = true;
            style.IncludeBorder = true;
            style.IncludePatterns = true;
            style.IncludeProtection = true;
        }

        public static void AddErrorStyle(this Workbook workbook)
        {
            foreach (Style item in workbook.Styles)
                if (item.Name == "Error")
                    return;

            var style = workbook.Styles.Add("Error");
            style.Interior.Color = Color.Red;
            style.Font.Color = Color.White;
            style.IncludeNumber = true;
            style.IncludeFont = true;
            style.IncludeAlignment = true;
            style.IncludeBorder = true;
            style.IncludePatterns = true;
            style.IncludeProtection = true;
        }

        #endregion

        #region worksheet
        public static void ExportTxt(this Worksheet worksheet, string path)
        {
            var file = Path.Combine(path, worksheet.Name + ".txt");
            if (File.Exists(file))
                File.Delete(file);

            using (var sw = File.CreateText(file))
            {
                if (worksheet.UsedRange.Count == 1)
                {
                    var rowCount = worksheet.UsedRange.Rows.Count;
                    var colCount = worksheet.UsedRange.Columns.Count;
                    object data = worksheet.UsedRange.Formula;
                    if (data != null)
                        for (var i = 1; i <= rowCount; i++)
                        {
                            for (var j = 1; j <= colCount; j++)
                            {
                                var value = data == null ? "" : data.ToString();
                                sw.Write(value);
                                sw.Write("\t");
                            }

                            sw.Write(Environment.NewLine);
                        }
                }
                else
                {
                    var rowCount = worksheet.UsedRange.Rows.Count;
                    var colCount = worksheet.UsedRange.Columns.Count;
                    object[,] data = worksheet.UsedRange.Formula;
                    if (data != null)
                        for (var i = 1; i <= rowCount; i++)
                        {
                            for (var j = 1; j <= colCount; j++)
                            {
                                var value = data[i, j] == null ? "" : data[i, j].ToString();
                                sw.Write(value);
                                sw.Write("\t");
                            }

                            sw.Write(Environment.NewLine);
                        }
                }
            }
        }

        public static int GetColumnIndexByHeader(this Worksheet worksheet, string firstHeader, string headerGroupName,
            out int headerRowNumber)
        {
            var startColNumber = worksheet.UsedRange.Column;
            var startRowNumber = worksheet.UsedRange.Row;
            var stopColNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            var stopRowNumber = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            var dataArray = worksheet.GetRange(startRowNumber, startColNumber, stopRowNumber, stopColNumber).Value2;
            headerRowNumber = -1;

            var rowNum = stopRowNumber > 10 ? 10 : stopRowNumber;
            var colNum = stopColNumber > 10 ? 10 : stopColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                {
                    var header = dataArray[i, j] == null ? "" : dataArray[i, j].ToString().Trim();
                    if (header.Equals(firstHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        headerRowNumber = i;
                        break;
                    }
                }

            if (headerRowNumber != -1)
                for (var j = 1; j <= colNum; j++)
                {
                    var header = dataArray[headerRowNumber, j] == null
                        ? ""
                        : dataArray[headerRowNumber, j].ToString().Trim();
                    if (header.Equals(headerGroupName, StringComparison.OrdinalIgnoreCase))
                        return j;
                }

            return -1;
        }

        public static int GetColumnIndexByHeader(this Worksheet worksheet, string headerGroupName)
        {
            var startColNumber = worksheet.UsedRange.Column;
            var startRowNumber = worksheet.UsedRange.Row;
            var stopColNumber = worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1;
            var stopRowNumber = worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1;
            var dataArray = worksheet.GetRange(startRowNumber, startColNumber, stopRowNumber, stopColNumber).Value2;

            var rowNum = stopRowNumber > 10 ? 10 : stopRowNumber;
            var colNum = stopColNumber > 10 ? 10 : stopColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                {
                    var header = dataArray[i, j] == null ? "" : dataArray[i, j].ToString().Trim();
                    if (header.Equals(headerGroupName, StringComparison.OrdinalIgnoreCase))
                        return j;
                }

            return -1;
        }

        public static void SetHeaderStyle(this Worksheet worksheet)
        {
            var workbook = (Workbook)worksheet.Parent;
            workbook.AddHeaderStyle();
            worksheet.GetRange(worksheet.UsedRange.Row, worksheet.UsedRange.Column,
                    worksheet.UsedRange.Row, worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1).Style =
                "Header";
        }

        public static void SetOutline<T>(this Worksheet worksheet, List<T> list,
            Expression<Func<T, string>> lambdaExpression) where T : new()
        {
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance;
            var colNum = typeof(T).GetProperties(memberFlags).Count(p => p.DeclaringType == typeof(T));
            var first = typeof(T).GetProperties(memberFlags).First(p => p.DeclaringType == typeof(T));
            var last = typeof(T).GetProperties(memberFlags).Last(p => p.DeclaringType == typeof(T));
            var getter = lambdaExpression.Compile();

            worksheet.UsedRange.ClearOutline();
            var orderList = new List<T>();
            var rowNum = 1;
            foreach (var group in list.GroupBy(getter).ToList())
            {
                orderList.AddRange(group);

                #region add summary row
                var summaryRow = new T();
                var startRow = rowNum + 1;
                var endRow = rowNum + group.Count();
                rowNum += group.Count() + 1;
                var formula = "=COUNTA(" + GetAddress(startRow, colNum) + ":" +
                              GetAddress(endRow, colNum) + ")";
                first.SetValue(summaryRow, group.Select(getter).First(), null);
                last.SetValue(summaryRow, formula, null);
                orderList.Add(summaryRow);
                #endregion
            }

            var range = (Range)worksheet.Cells[1, 1];
            range.LoadFromCollection(orderList);
            worksheet.UsedRange.AutoOutline();
            worksheet.Outline.ShowLevels(2);
            worksheet.UsedRange.AutoFilter(1);
            worksheet.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        public static Range GetRange(this Worksheet worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            return worksheet.Range[worksheet.Cells[fromRow, fromCol], worksheet.Cells[toRow, toCol]];
        }

        public static void MergeColumn(this Worksheet worksheet, int colNum, bool printHeaders = true)
        {
            var array = worksheet.Range[worksheet.Cells[1, colNum],
                worksheet.Cells[worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1, colNum]].Value2;
            if (array.count == 1) return;

            var lastIndex = 0;
            var start = printHeaders ? 2 : 1;
            for (var i = start; i <= array.Length; i++)
            {
                for (var j = i + 1; j <= array.Length; j++)
                {
                    if (array[i, colNum] != array[j, colNum])
                    {
                        lastIndex = j - 1;
                        break;
                    }

                    if (j == array.Length)
                    {
                        lastIndex = j;
                        break;
                    }
                }

                if (lastIndex - i > 0)
                    worksheet.Range[worksheet.Cells[i, colNum], worksheet.Cells[lastIndex, colNum]].Merge();
            }
        }

        #endregion

        #region range

        //public static Range LoadFromCollectionInverse<T>(this Range range, IEnumerable<T> collection,
        //    bool printHeaders = true)
        //{
        //    return LoadFromCollection(range, collection, true, printHeaders);
        //}

        public static Range LoadFromCollection<T>(this Range range, IEnumerable<T> collection,
            bool printHeaders = true)
        {
            return LoadFromCollection(range, collection, false, printHeaders);
        }

        private static Range LoadFromCollection<T>(Range range, IEnumerable<T> collection, bool isInverse,
            bool isPrintHeaders)
        {
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance;
            var type = typeof(T);
            var members = new List<MemberInfo>();
            members.AddRange(type.GetProperties(memberFlags).Where(p => p.DeclaringType == typeof(T)).ToList());

            if (members.Count == 0)
                throw new ArgumentException("Parameter Members must have at least one Property. Length is zero");

            var enumerable = collection as T[] ?? collection.ToArray();
            var values = new object[isPrintHeaders ? enumerable.Length + 1 : enumerable.Length, members.Count];
            int col = 0, row = 0;
            if (members.Count > 0 && isPrintHeaders)
            {
                foreach (var t in members)
                {
                    string header;
                    if (t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() is
                        DescriptionAttribute)
                    {
                        var descriptionAttribute =
                            (DescriptionAttribute)t.GetCustomAttributes(typeof(DescriptionAttribute), false)
                                .FirstOrDefault();
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        if (t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() is
                            DisplayNameAttribute)
                        {
                            var displayNameAttribute =
                                (DisplayNameAttribute)t.GetCustomAttributes(typeof(DisplayNameAttribute), false)
                                    .FirstOrDefault();
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            header = t.Name.Replace('_', ' ');
                        }
                    }

                    values[row, col++] = header;
                }

                row++;
            }

            if (!enumerable.Any() && (members.Count == 0 || isPrintHeaders == false)) return null;

            foreach (var item in enumerable)
            {
                col = 0;
                if (item is string || item is decimal || item is DateTime || item.GetType().IsPrimitive)
                {
                    values[row, col++] = item;
                }
                else
                {
                    var index = 0;
                    foreach (var t in members)
                        if (t is PropertyInfo)
                        {
                            var propertyInfo = (PropertyInfo)t;
                            var parameters = propertyInfo.GetIndexParameters();
                            if (!parameters.Any())
                                values[row, col++] = propertyInfo.GetValue(item, null);
                            else
                                values[row, col++] = propertyInfo.GetValue(item, new object[] { index++ });
                        }
                        else if (t is FieldInfo)
                        {
                            var fieldInfo = (FieldInfo)t;
                            values[row, col++] = fieldInfo.GetValue(item);
                        }
                        else if (t is MethodInfo)
                        {
                            var methodInfo = (MethodInfo)t;
                            values[row, col++] = methodInfo.Invoke(item, null);
                        }
                }

                row++;
            }

            var fromRow = range.Row;
            var fromCol = range.Column;
            Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                fromCol + values.GetLength(1) - 1);
            myRange.Value = values;
            return myRange;
        }

        public static Range LoadFromArrayInverse(this Range range, IEnumerable<object[]> data,
            List<string> headerList = null)
        {
            return LoadFromArrays(range, data, true, headerList);
        }

        public static Range LoadFromArrays(this Range range, IEnumerable<object[]> data, List<string> headerList = null)
        {
            return LoadFromArrays(range, data, false, headerList);
        }

        private static Range LoadFromArrays(Range range, IEnumerable<object[]> data, bool isInverse,
            List<string> headerList = null)
        {
            if (data == null) throw new ArgumentNullException();

            var isPrintHeaders = !(headerList == null || headerList.Count == 0);

            var rowArray = new List<object[]>();
            var maxRow = 0;
            var enumerable = data as object[][] ?? data.ToArray();
            foreach (var item in enumerable)
            {
                rowArray.Add(item);
                if (maxRow < item.Length) maxRow = item.Length;
            }

            maxRow = isPrintHeaders ? maxRow + 1 : maxRow;
            var minCol = headerList == null || headerList.Count == 0
                ? rowArray.Count
                : Math.Min(rowArray.Count, headerList.Count);

            if (rowArray.Count == 0) return null;

            if (isInverse)
            {
                var values = new object[maxRow, minCol];
                int col = 0;
                int row = 0;
                if (maxRow > 0 && isPrintHeaders)
                    for (var i = 0; i < minCol; i++)
                        values[row, col++] = headerList[i];

                col = 0;
                foreach (var item in enumerable)
                {
                    row = 0;
                    foreach (var t in item) values[row++, col] = t;
                    col++;
                }
                var fromRow = range.Row;
                var fromCol = range.Column;
                Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                   fromCol + values.GetLength(1) - 1);
                myRange.Value = values;
                return myRange;
            }
            else
            {
                var values = new object[minCol, maxRow];
                int col = 0;
                int row = 0;
                if (maxRow > 0 && isPrintHeaders)
                    for (var i = 0; i < minCol; i++)
                        values[col++, row] = headerList[i];
                col = 0;
                foreach (var item in enumerable)
                {
                    row = 0;
                    foreach (var t in item) values[col, row++] = t;
                    col++;
                }
                var fromRow = range.Row;
                var fromCol = range.Column;
                Range myRange = range.Worksheet.GetRange(fromRow, fromCol, fromRow + values.GetLength(0) - 1,
                    fromCol + values.GetLength(1) - 1);
                myRange.Value = values;
                return myRange;
            }
        }

        public static void AddErrorStyle(this Range range)
        {
            range.Style = "Error";
        }

        public static void AddHeaderStyle(this Range range)
        {
            range.Style = "Header";
        }

        //public static string GetValue2((this Range range)
        //{
        //    if (!range.MergeCells)
        //        return range.Value2 ?? string.Empty;

        //    return range.MergeArea.Cells[1.1].Value2 ?? string.Empty;
        //}

        public static string Address(this Range range)
        {
            var startRow = range.Row;
            var stopRow = range.Row + range.Rows.Count - 1;
            var startCol = range.Column;
            var stopCol = range.Column + range.Columns.Count - 1;

            return "$" + GetColumnLetter(startCol) + "$" + startRow + ":" + "$" + GetColumnLetter(stopCol) + "$" +
                   stopRow;
        }

        public static string GetAddress(int row, int column, bool absolute = false)
        {
            if (row == 0 || column == 0) return "#REF!";

            if (absolute) return "$" + GetColumnLetter(column) + "$" + row;

            return GetColumnLetter(column) + row;
        }

        public static string GetColumnLetter(this Range range)
        {
            var startCol = range.Column;
            return GetColumnLetter(startCol);
        }

        private static string GetColumnLetter(int colNum)
        {
            if (colNum < 1) return "#REF!";

            var sCol = "";
            do
            {
                sCol = (char)('A' + (colNum - 1) % 26) + sCol;
                colNum = (colNum - (colNum - 1) % 26) / 26;
            } while (colNum > 0);

            return sCol;
        }

        #endregion
    }
}