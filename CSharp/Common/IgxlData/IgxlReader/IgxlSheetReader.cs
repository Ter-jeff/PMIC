using IgxlData.IgxlSheets;
using Ionic.Zip;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using Teradyne.Oasis.IGData.Utilities;
using BinTableSheet = IgxlData.IgxlSheets.BinTableSheet;
using PortMapSheet = IgxlData.IgxlSheets.PortMapSheet;

namespace IgxlData.IgxlReader
{
    public class IgxlSheetReader
    {
        private readonly Dictionary<string, Dictionary<string, SheetObjMap>> _igxlConfigDic =
            new Dictionary<string, Dictionary<string, SheetObjMap>>();

        public IgxlSheetReader()
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //var resourceNames = assembly.GetManifestResourceNames();
            //foreach (var resourceName in resourceNames)
            //{
            //    if (resourceName.Contains(".SheetClassMapping."))
            //    {
            //        var igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
            //        foreach (var sheetItemClass in igxlConfig.SheetItemClass)
            //        {
            //            var sheetName = sheetItemClass.sheetname;
            //            var sheetVersion = sheetItemClass.sheetversion;
            //            Dictionary<string, SheetObjMap> dic = new Dictionary<string, SheetObjMap>();
            //            if (!dic.ContainsKey(sheetVersion))
            //            {
            //                dic.Add(sheetVersion, sheetItemClass);
            //                if (!_igxlConfigDic.ContainsKey(sheetName))
            //                    _igxlConfigDic.Add(sheetName, dic);
            //                else
            //                {
            //                    if (!_igxlConfigDic[sheetName].ContainsKey(sheetVersion))
            //                        _igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
            //                }
            //            }
            //        }
            //    }
            //}

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase)
                .Replace("file:\\", "");
            var files = Directory.GetFiles(Path.Combine(exePath, "IGDataXML\\SheetClassMapping"));
            foreach (var file in files)
                if (file.EndsWith("_ultraflex.xml", StringComparison.CurrentCultureIgnoreCase))
                {
                    var igxlConfig = LoadConfig(File.OpenRead(file));
                    foreach (var sheetItemClass in igxlConfig.SheetItemClass)
                    {
                        var sheetName = sheetItemClass.sheetname;
                        var sheetVersion = sheetItemClass.sheetversion;
                        var dic = new Dictionary<string, SheetObjMap>();
                        if (!dic.ContainsKey(sheetVersion))
                        {
                            dic.Add(sheetVersion, sheetItemClass);
                            if (!_igxlConfigDic.ContainsKey(sheetName))
                            {
                                _igxlConfigDic.Add(sheetName, dic);
                            }
                            else
                            {
                                if (!_igxlConfigDic[sheetName].ContainsKey(sheetVersion))
                                    _igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
                            }
                        }
                    }
                }
        }

        private IGXL LoadConfig(Stream sr)
        {
            IGXL result;
            try
            {
                var xs = new XmlSerializer(typeof(IGXL));
                var sysData = (IGXL)xs.Deserialize(sr);
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            return result;
        }

        public string GetCellText(string[] arr, int i)
        {
            if (i < arr.Length)
                return arr[i];
            return "";
        }

        public string GetCellText(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet.Cells[row, column] != null)
            {
                if (sheet.Cells[row, column].Formula != string.Empty)
                    return "=" + sheet.Cells[row, column].Formula;

                if (sheet.Cells[row, column].Value != null)
                    if (sheet.Cells[row, column].Value is double ||
                        sheet.Cells[row, column].Value is bool)
                        return sheet.Cells[row, column].Value.ToString();
                return sheet.Cells[row, column].Text;
            }

            return "";
        }

        public string GetCellText(Worksheet sheet, int row, int column)
        {
            if (sheet.Cells[row, column] != null)
            {
                if (sheet.Cells[row, column].Formula != string.Empty)
                    return "=" + sheet.Cells[row, column].Formula;

                if (sheet.Cells[row, column].Value != null)
                    if (sheet.Cells[row, column].Value is double ||
                        sheet.Cells[row, column].Value is bool)
                        return sheet.Cells[row, column].Value.ToString();

                if (sheet.Cells[row, column].Value != null &&
                    sheet.Cells[row, column].Value != sheet.Cells[row, column].Text)
                {
                }

                return sheet.Cells[row, column].Text;
            }

            return "";
        }

        public string GetMergeCellValue(ExcelWorksheet sheet, int row, int column)
        {
            var range = sheet.MergedCells[row, column];
            var mergeCellValue = range == null
                ? GetCellText(sheet, row, column)
                : GetCellText(sheet, new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column);
            return mergeCellValue;
        }

        public static string FormatStringForCompare(string pString)
        {
            var lStrResult = pString.Trim();

            lStrResult = ReplaceDoubleBlank(lStrResult);

            lStrResult = lStrResult.Replace(" ", "_");

            lStrResult = lStrResult.ToUpper();

            return lStrResult;
        }

        public static string ReplaceDoubleBlank(string pString)
        {
            var lStrResult = pString;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        public static bool IsLiked(string pStrInput, string pStrPatten)
        {
            if (pStrPatten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@".+", StringComparison.Ordinal) >= 0)
                return Regex.IsMatch(FormatStringForCompare(pStrInput), "^" + FormatStringForCompare(pStrPatten) + "$");
            return FormatStringForCompare(pStrInput) == FormatStringForCompare(pStrPatten);
        }

        public IgxlSheet CreateIgxlSheet(ExcelWorksheet sheet)
        {
            var sheetType = GetIgxlSheetType(GetCellText(sheet, 1, 1));
            return GetIgxlSheet(sheet, sheetType);
        }

        public IgxlSheet GetIgxlSheet(ExcelWorksheet sheet, SheetTypes sheetType)
        {
            switch (sheetType)
            {
                case SheetTypes.DTFlowtableSheet:
                    return new ReadFlowSheet().GetSheet(sheet);
                case SheetTypes.DTTestInstancesSheet:
                    return new ReadInstanceSheet().GetSheet(sheet);
                case SheetTypes.DTLevelSheet:
                    return new ReadLevelSheet().GetSheet(sheet);
                case SheetTypes.DTPatternSetSheet:
                    return new ReadPatSetSheet().GetSheet(sheet);
                case SheetTypes.DTTimesetBasicSheet:
                    return new ReadTimeSetSheet().GetSheet(sheet);
                case SheetTypes.DTChanMap:
                    return new ReadChanMapSheet().GetSheet(sheet);
                case SheetTypes.DTJobListSheet:
                    return new ReadJobListSheet().GetSheet(sheet);
                case SheetTypes.DTGlobalSpecSheet:
                    return new ReadGlobalSpecSheet().GetSheet(sheet);
                case SheetTypes.DTACSpecSheet:
                    return new ReadAcSpecSheet().GetSheet(sheet);
                case SheetTypes.DTDCSpecSheet:
                    return new ReadDcSpecSheet().GetSheet(sheet);
                case SheetTypes.DTPinMap:
                    return new ReadPinMapSheet().GetSheet(sheet);
                case SheetTypes.DTPortMapSheet:
                    return GetPortMapSheet(sheet);
                case SheetTypes.DTUnknown:
                    break;
            }

            return null;
        }

        public SheetTypes GetIgxlSheetTypeByFile(string file)
        {
            var firstLine = ReadFirstLine(file);
            return GetIgxlSheetType(firstLine);
        }

        public IgxlSheet CreateIgxlSheet(string file)
        {
            var firstLine = ReadFirstLine(file);
            var sheetType = GetIgxlSheetType(firstLine);
            var sheet = ConvertTxtToExcelSheet(file);
            return GetIgxlSheet(sheet, sheetType);
        }

        public IgxlSheet CreateIgxlSheet(string sheetName, Stream stream, SheetTypes sheetType)
        {
            var sheet = ConvertStreamToExcelSheet(sheetName, stream);
            return GetIgxlSheet(sheet, sheetType);
        }

        public ExcelWorksheet ConvertStreamToExcelSheet(string sheetName, Stream stream)
        {
            var excelPackage = new ExcelPackage();
            sheetName = sheetName.Replace("%20", " ");
            var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            var index = 0;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    index++;
                    if (line != null)
                    {
                        var arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
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

        public ExcelWorksheet ConvertTxtToExcelSheet(string fileName)
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
                        var arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
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

        public PortMapSheet GetPortMapSheet(ExcelWorksheet worksheet)
        {
            var firstLine = GetCellText(worksheet, 1, 1);
            var version = GetIgxlSheetVersion(firstLine);
            var key = IgxlSheetNameList.PortMap;
            SheetObjMap sheetObjMap = null;
            if (_igxlConfigDic.ContainsKey(key))
                if (_igxlConfigDic[key].ContainsKey(version))
                    sheetObjMap = _igxlConfigDic[key][version];
            var readPortMapSheet = new ReadPortMapSheet(sheetObjMap);
            return readPortMapSheet.GetSheet(worksheet);
        }

        public PortMapSheet GetPortMapSheet(Worksheet worksheet)
        {
            var firstLine = GetCellText(worksheet, 1, 1);
            var version = GetIgxlSheetVersion(firstLine);
            var key = IgxlSheetNameList.PortMap;
            SheetObjMap sheetObjMap = null;
            if (_igxlConfigDic.ContainsKey(key))
                if (_igxlConfigDic[key].ContainsKey(version))
                    sheetObjMap = _igxlConfigDic[key][version];
            var readPortMapSheet = new ReadPortMapSheet(sheetObjMap);
            return readPortMapSheet.GetSheet(worksheet);
        }

        #region read Excel

        public List<IgxlSheet> GetIgxlSheets(ExcelWorkbook workbook, SheetTypes type)
        {
            var igxlSheets = new List<IgxlSheet>();
            foreach (var sheet in workbook.Worksheets)
                if (GetIgxlSheetType(GetCellText(sheet, 1, 1)) == type)
                    igxlSheets.Add(CreateIgxlSheet(sheet));
            return igxlSheets;
        }

        public List<string> GetIgxlSheetList(ExcelWorkbook workbook, SheetTypes type)
        {
            var sheets = workbook.Worksheets.Where(x => GetIgxlSheetType(GetCellText(x, 1, 1)) == type)
                .Select(x => x.Name).ToList();
            return sheets;
        }

        #endregion

        #region read *.txt
        public List<IgxlSheet> GetIgxlSheetsByFolder(string path, SheetTypes type)
        {
            var igxlSheets = new List<IgxlSheet>();
            var files = Directory.GetFiles(path);
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file);
                if (!extension.Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                var firstLine = ReadFirstLine(file);
                var sheetType = GetIgxlSheetType(firstLine);
                if (sheetType == type)
                    igxlSheets.Add(CreateIgxlSheet(file));
            }

            return igxlSheets;
        }

        public Dictionary<string, SheetTypes> GetSheetTypeDic(string path)
        {
            var dic = new Dictionary<string, SheetTypes>();
            var files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories); //Modified by terry
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file);
                if (!extension.Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                var firstLine = ReadFirstLine(file);
                var type = GetIgxlSheetType(firstLine);
                if (!dic.ContainsKey(file))
                    dic.Add(file, type);
            }

            return dic;
        }

        public List<string> GetSheetByType(string path, SheetTypes type)
        {
            var dic = GetSheetTypeDic(path);
            return dic.ToList().Where(x => x.Value == type).Select(x => x.Key).ToList();
        }

        public bool IsSheetType(SheetTypes type, string file)
        {
            if (Regex.IsMatch(ReadFirstLine(file), type.ToString(), RegexOptions.IgnoreCase))
                return true;
            return false;
        }

        public string GetIgxlSheetVersion(string text)
        {
            if (Regex.IsMatch(text, @"version=(?<value>\d+([.]\d+)?)", RegexOptions.IgnoreCase))
            {
                var match = Regex.Match(text, @"version=(?<value>\d+([.]\d+)?)", RegexOptions.IgnoreCase);
                return match.Groups["value"].ToString();
            }

            return "";
        }

        public SheetTypes GetIgxlSheetType(string text)
        {
            if (string.IsNullOrEmpty(text))
                return SheetTypes.DTUnknown;

            if (Regex.IsMatch(text, SheetTypes.DTFlowtableSheet.ToString()))
                return SheetTypes.DTFlowtableSheet;
            if (Regex.IsMatch(text, SheetTypes.DTTestInstancesSheet.ToString()))
                return SheetTypes.DTTestInstancesSheet;
            if (Regex.IsMatch(text, SheetTypes.DTDCSpecSheet.ToString()))
                return SheetTypes.DTDCSpecSheet;
            if (Regex.IsMatch(text, SheetTypes.DTACSpecSheet.ToString()))
                return SheetTypes.DTACSpecSheet;
            if (Regex.IsMatch(text, SheetTypes.DTLevelSheet.ToString()))
                return SheetTypes.DTLevelSheet;
            if (Regex.IsMatch(text, SheetTypes.DTGlobalSpecSheet.ToString()))
                return SheetTypes.DTGlobalSpecSheet;
            if (Regex.IsMatch(text, SheetTypes.DTTimesetBasicSheet.ToString()))
                return SheetTypes.DTTimesetBasicSheet;
            if (Regex.IsMatch(text, SheetTypes.DTBintablesSheet.ToString()))
                return SheetTypes.DTBintablesSheet;
            if (Regex.IsMatch(text, SheetTypes.DTChanMap.ToString()))
                return SheetTypes.DTChanMap;
            if (Regex.IsMatch(text, SheetTypes.DTCharacterizationSheet.ToString()))
                return SheetTypes.DTCharacterizationSheet;
            if (Regex.IsMatch(text, SheetTypes.DTJobListSheet.ToString()))
                return SheetTypes.DTJobListSheet;
            if (Regex.IsMatch(text, SheetTypes.DTPatternSetSheet.ToString()))
                return SheetTypes.DTPatternSetSheet;
            if (Regex.IsMatch(text, SheetTypes.DTPatternSubroutineSheet.ToString()))
                return SheetTypes.DTPatternSubroutineSheet;
            if (Regex.IsMatch(text, SheetTypes.DTPinMap.ToString()))
                return SheetTypes.DTPinMap;
            if (Regex.IsMatch(text, SheetTypes.DTPortMapSheet.ToString()))
                return SheetTypes.DTPortMapSheet;

            return SheetTypes.DTUnknown;
        }

        private string ReadFirstLine(string file)
        {
            using (var sw = new StreamReader(file))
            {
                return sw.ReadLine() ?? "";
            }
        }

        public List<string> GetEnables(string igxl)
        {
            var enables = new List<string>();
            using (var zip = new ZipFile(igxl))
            {
                var entries = zip.Entries.ToList();
                foreach (var entry in entries)
                {
                    var sheetName = Path.GetFileNameWithoutExtension(entry.FileName);
                    if (sheetName != null)
                    {
                        var stream = entry.OpenReader();
                        var firstLine = "";
                        using (var sr = new StreamReader(stream))
                        {
                            firstLine = sr.ReadLine();
                        }

                        stream = entry.OpenReader();
                        var type = GetIgxlSheetType(firstLine);
                        if (type == SheetTypes.DTFlowtableSheet)
                            enables.AddRange(new ReadFlowSheet().GetEnables(stream, sheetName));
                    }
                }
            }

            enables = enables.SelectMany(x => Regex.Split(x, @"\&|\||!|\(|\)"))
                .Select(x => x.Trim()).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            return enables.OrderBy(x => x).ToList();
        }

        public List<string> GetJobs(string igxl)
        {
            var jobs = new List<string>();
            using (var zip = new ZipFile(igxl))
            {
                var entries = zip.Entries.ToList();
                foreach (var entry in entries)
                {
                    var sheetName = Path.GetFileNameWithoutExtension(entry.FileName);
                    if (sheetName != null)
                    {
                        var stream = entry.OpenReader();
                        var firstLine = "";
                        using (var sr = new StreamReader(stream))
                        {
                            firstLine = sr.ReadLine();
                        }

                        stream = entry.OpenReader();
                        var type = GetIgxlSheetType(firstLine);
                        if (type == SheetTypes.DTJobListSheet)
                            jobs.AddRange(new ReadJobListSheet().GetJobs(stream, sheetName));
                    }
                }
            }

            return jobs.Distinct().ToList();
        }

        public SubFlowSheet GetSubFlowSheet(string name, Stream stream)
        {
            return new ReadFlowSheet().GetSheet(ConvertStreamToExcelSheet(name, stream));
        }

        public JobListSheet GetJobListSheet(string name, Stream stream)
        {
            return new ReadJobListSheet().GetSheet(ConvertStreamToExcelSheet(name, stream));
        }

        public BinTableSheet GetBinTableSheet(string name, Stream stream)
        {
            return new ReadBinTableSheet().GetSheet(ConvertStreamToExcelSheet(name, stream));
        }

        public IgxlSheet GetIgxlSheet(Stream stream, string sheetName, SheetTypes sheetType)
        {
            switch (sheetType)
            {
                case SheetTypes.DTFlowtableSheet:
                    {
                        return new ReadFlowSheet().GetSheet(stream, sheetName);
                    }
                case SheetTypes.DTTestInstancesSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadInstanceSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTLevelSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadLevelSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTPatternSetSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadPatSetSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTPatternSubroutineSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadPatSubroutineSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTTimesetBasicSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadTimeSetSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTBintablesSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadBinTableSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTChanMap:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadChanMapSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTJobListSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadJobListSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTGlobalSpecSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadGlobalSpecSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTACSpecSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadAcSpecSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTDCSpecSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadDcSpecSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTPinMap:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return new ReadPinMapSheet().GetSheet(sheet);
                    }
                case SheetTypes.DTPortMapSheet:
                    {
                        var sheet = ConvertStreamToExcelSheet(sheetName, stream);
                        return GetPortMapSheet(sheet);
                    }
                case SheetTypes.DTUnknown:
                    break;
            }
            return null;
        }
        #endregion

        public List<string> GetSites(string igxl)
        {
            var terChanMap = "";
            var sites = new List<string>();
            if (!File.Exists(igxl))
                return sites;
            using (var zip = new ZipFile(igxl))
            {
                var entries = zip.Entries.ToList();
                foreach (var entry in entries)
                {
                    var sheetName = Path.GetFileNameWithoutExtension(entry.FileName);
                    if (sheetName == "tl_WorkBookProperties_")
                    {
                        var stream = entry.OpenReader();
                        var lines = new List<string>();
                        using (var sr = new StreamReader(stream))
                        {
                            while (!sr.EndOfStream)
                            {
                                var line = sr.ReadLine();
                                lines.Add(line);
                            }
                        }
                        terChanMap = GetTerChanMap(lines);
                        break;
                    }
                }
            }

            if (!string.IsNullOrEmpty(terChanMap))
            {
                using (var zip = new ZipFile(igxl))
                {
                    var entries = zip.Entries.ToList();
                    foreach (var entry in entries)
                    {
                        var sheetName = Path.GetFileNameWithoutExtension(entry.FileName);
                        if (sheetName == terChanMap)
                        {
                            var stream = entry.OpenReader();
                            string firstLine = "";
                            using (var sr = new StreamReader(stream))
                                firstLine = sr.ReadLine();

                            stream = entry.OpenReader();
                            SheetTypes type = GetIgxlSheetType(firstLine);
                            if (type == SheetTypes.DTChanMap)
                                sites.AddRange(new ReadChanMapSheet().GetSites(stream, sheetName));
                        }
                    }
                }
            }

            return sites.ToList();
        }

        private string GetTerChanMap(List<string> lines)
        {
            foreach (var line in lines)
            {
                var arr = line.Split('\t');
                if (arr.First().Equals("terChanMap", StringComparison.CurrentCultureIgnoreCase))
                    return arr.Last();
            }
            return "";
        }
    }
}
