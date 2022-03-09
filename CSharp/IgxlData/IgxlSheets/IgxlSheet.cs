using IgxlData.IgxlBase;
using OfficeOpenXml;
using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Teradyne.Oasis.IGData.Utilities;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    public abstract class IgxlSheet : IgxlItem
    {
        #region Field
        protected StreamWriter IgxlWriter;
        protected string IgxlSheetContext;
        #endregion

        #region Property
        public string SheetName { get; set; }
        public string IgxlSheetName { get; set; }
        public string JobName { get; set; }
        #endregion

        #region Constructor

        protected IgxlSheet(Worksheet sheet)
        {
            IgxlSheetContext = sheet.Cells[1, 1].Text;
            SheetName = sheet.Name;
        }

        protected IgxlSheet(ExcelWorksheet sheet)
        {
            IgxlSheetContext = sheet.Cells[1, 1].Text;
            SheetName = sheet.Name;
        }

        protected IgxlSheet(string sheetName)
        {
            SheetName = sheetName;
        }
        #endregion

        #region Member function

        protected void GetStreamWriter(string fileName)
        {
            IgxlWriter = new StreamWriter(fileName);
        }

        protected abstract void WriteHeader();

        protected abstract void WriteColumnsHeader();

        protected abstract void WriteRows();

        public abstract void Write(string fileName, string version = "");

        protected void CloseStreamWriter()
        {
            if (IgxlWriter != null)
                IgxlWriter.Close();
        }

        public string GetVersion()
        {
            if (!string.IsNullOrEmpty(IgxlSheetContext))
            {
                int startIndex = IgxlSheetContext.IndexOf("version=", StringComparison.Ordinal);
                if (startIndex != -1)
                {
                    string text = IgxlSheetContext.Substring(startIndex + 8);
                    int stopIndex = text.IndexOf(":", StringComparison.Ordinal);
                    return text.Substring(0, stopIndex);
                }
            }
            return "";
        }

        protected int GetIndexFrom(SheetInfo sheetInfo, string name)
        {
            if (sheetInfo.Field != null)
            {
                foreach (var item in sheetInfo.Field)
                {
                    if (item.fieldName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.columnIndex;
                }
            }

            if (sheetInfo.Columns.Column != null)
            {
                foreach (var item in sheetInfo.Columns.Column)
                {
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (var item in sheetInfo.Columns.Variant)
                {
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;
                }
            }

            if (sheetInfo.Columns.RelativeColumn != null)
            {
                foreach (var item in sheetInfo.Columns.RelativeColumn)
                {
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;
                }
            }
            return -1;
        }

        protected int GetMaxCount(SheetInfo sheetInfo)
        {
            int max = -1;
            if (sheetInfo.Field != null)
            {
                int fieldMax = sheetInfo.Field.Max(x => x.columnIndex);
                if (fieldMax > max) max = fieldMax;
            }

            if (sheetInfo.Columns.Column != null)
            {
                foreach (var item in sheetInfo.Columns.Column)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                    {
                        if (item.Column1 != null)
                        {
                            foreach (var column1 in item.Column1)
                                max = Math.Max(max, column1.indexFrom);
                        }
                    }
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (var item in sheetInfo.Columns.Variant)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                    {
                        foreach (var column1 in item.Column1)
                            max = Math.Max(max, column1.indexFrom);
                    }
                }
            }

            if (sheetInfo.Columns.RelativeColumn != null)
            {
                int relativeColumnMax = sheetInfo.Columns.RelativeColumn.Max(x => x.indexFrom);
                if (relativeColumnMax > max) max = relativeColumnMax;
            }

            return max + 1;
        }

        protected Dictionary<string, Dictionary<string, SheetInfo>> GetIgxlSheetsVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            var igxlConfigDic = new Dictionary<string, Dictionary<string, SheetInfo>>();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.Contains(".IGXLSheetsVersion."))
                {
                    var xs = new XmlSerializer(typeof(IGXLVersion));
                    var igxlConfig = (IGXLVersion)xs.Deserialize(assembly.GetManifestResourceStream(resourceName));
                    foreach (var sheetItemClass in igxlConfig.Sheets)
                    {
                        var sheetName = sheetItemClass.sheetName;
                        var sheetVersion = sheetItemClass.sheetVersion;
                        var dic = new Dictionary<string, SheetInfo>();
                        if (!dic.ContainsKey(sheetVersion))
                        {
                            dic.Add(sheetVersion, sheetItemClass);
                            if (!igxlConfigDic.ContainsKey(sheetName))
                                igxlConfigDic.Add(sheetName, dic);
                            else
                            {
                                if (!igxlConfigDic[sheetName].ContainsKey(sheetVersion))
                                    igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
                            }

                        }
                    }
                }
            }
            return igxlConfigDic;
        }

        protected static void SetField(SheetInfo igxlSheetsVersion, int i, string[] arr)
        {
            if (igxlSheetsVersion.Field != null)
            {
                foreach (var item in igxlSheetsVersion.Field)
                {
                    if (item.rowIndex == i)
                        arr[item.columnIndex] = item.fieldName;
                }
            }
        }

        protected static void SetColumns(SheetInfo igxlSheetsVersion, int i, string[] arr)
        {
            if (igxlSheetsVersion.Columns.Column != null)
            {
                foreach (var item in igxlSheetsVersion.Columns.Column)
                {
                    if (item.rowIndex == i)
                        arr[item.indexFrom] = item.columnName;
                    if (item.Column1 != null)
                    {
                        foreach (var column1 in item.Column1)
                        {
                            if (column1.rowIndex == i)
                                arr[column1.indexFrom] = column1.columnName;
                        }
                    }
                }
            }
        }

        protected static void WriteHeader(string[] arr, StreamWriter sw)
        {
            if (arr.Any(x => !string.IsNullOrEmpty(x)))
                sw.WriteLine(string.Join("\t", arr).TrimEnd('\t'));
            else
                sw.WriteLine('\t');
        }

        protected static void SetRelativeColumn(SheetInfo igxlSheetsVersion, int i, string[] arr, int relativeColumnIndex)
        {
            if (igxlSheetsVersion.Columns.RelativeColumn != null)
            {
                foreach (var item in igxlSheetsVersion.Columns.RelativeColumn)
                {
                    if (item.indexFrom == item.indexTo && item.rowIndex == i)
                        arr[relativeColumnIndex - 1 + item.indexFrom] = item.columnName;
                }
            }
        }

        protected int GetIndexFrom(SheetInfo sheetInfo, string name, string subString)
        {
            if (sheetInfo.Columns.Column != null)
            {
                foreach (var item in sheetInfo.Columns.Column)
                {
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(subString))
                            return item.indexFrom;

                        if (item.Column1 != null)
                        {
                            foreach (var column1 in item.Column1)
                            {
                                if (column1.columnName.Equals(subString, StringComparison.CurrentCultureIgnoreCase))
                                    return column1.indexFrom;
                            }
                        }
                    }
                }
            }

            if (sheetInfo.Columns.Variant != null)
            {
                foreach (var item in sheetInfo.Columns.Variant)
                {
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(subString))
                            return item.indexFrom;

                        if (item.Column1 != null)
                        {
                            foreach (var column1 in item.Column1)
                            {
                                if (column1.columnName.Equals(subString, StringComparison.CurrentCultureIgnoreCase))
                                    return column1.indexFrom;
                            }
                        }
                    }
                }
            }

            return -1;
        }
        #endregion
    }
}