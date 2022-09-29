using IgxlData.IgxlBase;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    [DebuggerDisplay("{SheetName}")]
    public abstract class IgxlSheet
    {
        protected string IgxlSheetContext;
        protected StreamWriter IgxlWriter;

        public string SheetName { get; set; }
        public string IgxlSheetName { get; set; }
        public string JobName { get; set; }

        protected void GetStreamWriter(string fileName)
        {
            IgxlWriter = new StreamWriter(fileName);
        }

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
                var startIndex = IgxlSheetContext.IndexOf("version=", StringComparison.Ordinal);
                if (startIndex != -1)
                {
                    var text = IgxlSheetContext.Substring(startIndex + 8);
                    var stopIndex = text.IndexOf(":", StringComparison.Ordinal);
                    return text.Substring(0, stopIndex);
                }
            }

            return "";
        }

        protected int GetIndexFrom(SheetInfo sheetInfo, string name)
        {
            if (sheetInfo.Field != null)
                foreach (var item in sheetInfo.Field)
                    if (item.fieldName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.columnIndex;

            if (sheetInfo.Columns.Column != null)
                foreach (var item in sheetInfo.Columns.Column)
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;

            if (sheetInfo.Columns.Variant != null)
                foreach (var item in sheetInfo.Columns.Variant)
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;

            if (sheetInfo.Columns.RelativeColumn != null)
                foreach (var item in sheetInfo.Columns.RelativeColumn)
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        return item.indexFrom;
            return -1;
        }

        protected int GetMaxCount(SheetInfo sheetInfo)
        {
            var max = -1;
            if (sheetInfo.Field != null)
            {
                var fieldMax = sheetInfo.Field.Max(x => x.columnIndex);
                if (fieldMax > max) max = fieldMax;
            }

            if (sheetInfo.Columns.Column != null)
                foreach (var item in sheetInfo.Columns.Column)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                        if (item.Column1 != null)
                            foreach (var column1 in item.Column1)
                                max = Math.Max(max, column1.indexFrom);
                }

            if (sheetInfo.Columns.Variant != null)
                foreach (var item in sheetInfo.Columns.Variant)
                {
                    max = Math.Max(max, item.indexFrom);
                    if (item.Column1 != null)
                        foreach (var column1 in item.Column1)
                            max = Math.Max(max, column1.indexFrom);
                }

            if (sheetInfo.Columns.RelativeColumn != null)
            {
                var relativeColumnMax = sheetInfo.Columns.RelativeColumn.Max(x => x.indexFrom);
                if (relativeColumnMax > max) max = relativeColumnMax;
            }

            return max + 1;
        }

        protected Dictionary<string, Dictionary<string, SheetInfo>> GetIgxlSheetsVersion()
        {
            var igxlConfigDic = new Dictionary<string, Dictionary<string, SheetInfo>>();

            //var assembly = Assembly.GetExecutingAssembly();
            //var resourceNames = assembly.GetManifestResourceNames();
            //foreach (var resourceName in resourceNames)
            //{
            //    if (resourceName.Contains(".IGXLSheetsVersion."))
            //    {
            //        var xs = new XmlSerializer(typeof(IGXLVersion));
            //        var igxlConfig = (IGXLVersion)xs.Deserialize(assembly.GetManifestResourceStream(resourceName));
            //        foreach (var sheetItemClass in igxlConfig.Sheets)
            //        {
            //            var sheetName = sheetItemClass.sheetName;
            //            var sheetVersion = sheetItemClass.sheetVersion;
            //            var dic = new Dictionary<string, SheetInfo>();
            //            if (!dic.ContainsKey(sheetVersion))
            //            {
            //                dic.Add(sheetVersion, sheetItemClass);
            //                if (!igxlConfigDic.ContainsKey(sheetName))
            //                    igxlConfigDic.Add(sheetName, dic);
            //                else
            //                {
            //                    if (!igxlConfigDic[sheetName].ContainsKey(sheetVersion))
            //                        igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
            //                }

            //            }
            //        }
            //    }
            //}

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase)
                .Replace("file:\\", "");
            var files = Directory.GetFiles(Path.Combine(exePath, "IGDataXML\\IGXLSheetsVersion"));
            foreach (var file in files)
                if (file.EndsWith("_ultraflex.xml", StringComparison.CurrentCultureIgnoreCase))
                {
                    var xs = new XmlSerializer(typeof(IGXLVersion));
                    var igxlConfig = (IGXLVersion)xs.Deserialize(File.OpenRead(file));
                    foreach (var sheetItemClass in igxlConfig.Sheets)
                    {
                        var sheetName = sheetItemClass.sheetName;
                        var sheetVersion = sheetItemClass.sheetVersion;
                        var dic = new Dictionary<string, SheetInfo>();
                        if (!dic.ContainsKey(sheetVersion))
                        {
                            dic.Add(sheetVersion, sheetItemClass);
                            if (!igxlConfigDic.ContainsKey(sheetName))
                            {
                                igxlConfigDic.Add(sheetName, dic);
                            }
                            else
                            {
                                if (!igxlConfigDic[sheetName].ContainsKey(sheetVersion))
                                    igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
                            }
                        }
                    }
                }

            return igxlConfigDic;
        }

        protected void SetField(SheetInfo igxlSheetsVersion, int i, string[] arr)
        {
            if (igxlSheetsVersion.Field != null)
                foreach (var item in igxlSheetsVersion.Field)
                    if (item.rowIndex == i)
                        arr[item.columnIndex] = item.fieldName;
        }

        protected void SetColumns(SheetInfo igxlSheetsVersion, int i, string[] arr)
        {
            if (igxlSheetsVersion.Columns.Column != null)
                foreach (var item in igxlSheetsVersion.Columns.Column)
                {
                    if (item.rowIndex == i)
                        arr[item.indexFrom] = item.columnName;
                    if (item.Column1 != null)
                        foreach (var column1 in item.Column1)
                            if (column1.rowIndex == i)
                                arr[column1.indexFrom] = column1.columnName;
                }
        }

        protected void WriteHeader(string[] arr, StreamWriter sw)
        {
            if (arr.Any(x => !string.IsNullOrEmpty(x)))
                sw.WriteLine(string.Join("\t", arr).TrimEnd('\t'));
            else
                sw.WriteLine('\t');
        }

        protected void SetRelativeColumn(SheetInfo igxlSheetsVersion, int i, string[] arr, int relativeColumnIndex)
        {
            if (igxlSheetsVersion.Columns.RelativeColumn != null)
                foreach (var item in igxlSheetsVersion.Columns.RelativeColumn)
                    if (item.indexFrom == item.indexTo && item.rowIndex == i)
                        arr[relativeColumnIndex - 1 + item.indexFrom] = item.columnName;
        }

        protected int GetIndexFrom(SheetInfo sheetInfo, string name, string subString)
        {
            if (sheetInfo.Columns.Column != null)
                foreach (var item in sheetInfo.Columns.Column)
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(subString))
                            return item.indexFrom;

                        if (item.Column1 != null)
                            foreach (var column1 in item.Column1)
                                if (column1.columnName.Equals(subString, StringComparison.CurrentCultureIgnoreCase))
                                    return column1.indexFrom;
                    }

            if (sheetInfo.Columns.Variant != null)
                foreach (var item in sheetInfo.Columns.Variant)
                    if (item.columnName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(subString))
                            return item.indexFrom;

                        if (item.Column1 != null)
                            foreach (var column1 in item.Column1)
                                if (column1.columnName.Equals(subString, StringComparison.CurrentCultureIgnoreCase))
                                    return column1.indexFrom;
                    }

            return -1;
        }

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

        protected List<IgxlRow> AddBackUpRows(List<IgxlRow> rows)
        {
            var mainList = rows.Where(x => !x.IsBackup).ToList();
            var backupList = rows.Where(x => x.IsBackup).ToList();
            if (backupList.Any())
            {
                var type = rows.First().GetType();
                var empty = (IgxlRow)Activator.CreateInstance(type);
                for (int i = 0; i < 10; i++)
                { mainList.Add(empty); }
                mainList.AddRange(backupList);
            }
            return mainList;
        }
    }

}