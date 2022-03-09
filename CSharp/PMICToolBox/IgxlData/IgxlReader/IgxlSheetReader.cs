using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.IgxlWorkBooks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using PortMapSheet = IgxlData.IgxlSheets.PortMapSheet;

namespace IgxlData.IgxlReader
{
    public class IgxlSheetReader
    {
        private readonly Dictionary<string, Dictionary<string, SheetObjMap>> _igxlConfigDic = new Dictionary<string, Dictionary<string, SheetObjMap>>();

        public ExcelWorksheet ConvertWorksheetToExcelSheet(Worksheet worksheet)
        {
            return worksheet.ConvertToExcelSheet();
        }


        private IGXL LoadConfig(Stream sr)
        {
            IGXL result;
            try
            {
                var xs = new XmlSerializer(typeof(IGXL));
                IGXL sysData = (IGXL)xs.Deserialize(sr);
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return result;
        }

        public IgxlSheetReader()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.Contains(".SheetClassMapping."))
                {
                    var igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
                    foreach (var sheetItemClass in igxlConfig.SheetItemClass)
                    {
                        var sheetName = sheetItemClass.sheetname;
                        var sheetVersion = sheetItemClass.sheetversion;
                        Dictionary<string, SheetObjMap> dic = new Dictionary<string, SheetObjMap>();
                        if (!dic.ContainsKey(sheetVersion))
                        {
                            dic.Add(sheetVersion, sheetItemClass);
                            if (!_igxlConfigDic.ContainsKey(sheetName))
                                _igxlConfigDic.Add(sheetName, dic);
                            else
                            {
                                if (!_igxlConfigDic[sheetName].ContainsKey(sheetVersion))
                                    _igxlConfigDic[sheetName].Add(sheetVersion, sheetItemClass);
                            }

                        }
                    }
                }
            }
        }

        public string GetCellText(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet.Cells[row, column] != null)
            {
                if (sheet.Cells[row, column].Formula != string.Empty)
                    return "=" + sheet.Cells[row, column].Formula;

                if (sheet.Cells[row, column].Value != null)
                {
                    if (sheet.Cells[row, column].Value is double ||
                       sheet.Cells[row, column].Value is bool)
                        return sheet.Cells[row, column].Value.ToString();
                }

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
                {
                    if (sheet.Cells[row, column].Value is double ||
                       sheet.Cells[row, column].Value is bool)
                        return sheet.Cells[row, column].Value.ToString();
                }

                if (sheet.Cells[row, column].Value != null && sheet.Cells[row, column].Value != sheet.Cells[row, column].Text)
                {
                }

                return sheet.Cells[row, column].Text;
            }
            return "";
        }

        public string GetMergeCellValue(ExcelWorksheet sheet, int row, int column)
        {
            string range = sheet.MergedCells[row, column];
            string mergeCellValue = range == null ? GetCellText(sheet, row, column) :
                GetCellText(sheet, (new ExcelAddress(range).Start.Row), (new ExcelAddress(range).Start.Column));
            return mergeCellValue;
        }

        public static string FormatStringForCompare(string pString)
        {
            string lStrResult = pString.Trim();

            lStrResult = ReplaceDoubleBlank(lStrResult);

            lStrResult = lStrResult.Replace(" ", "_");

            lStrResult = lStrResult.ToUpper();

            return lStrResult;
        }

        public static string ReplaceDoubleBlank(string pString)
        {
            string lStrResult = pString;
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
            {
                return Regex.IsMatch(FormatStringForCompare(pStrInput), "^" + FormatStringForCompare(pStrPatten) + "$");
            }
            return FormatStringForCompare(pStrInput) == FormatStringForCompare(pStrPatten);

        }

        public IgxlSheet CreateIgxlSheet(Worksheet sheet)
        {
            var sheetType = GetIgxlSheetType(GetCellText(sheet, 1, 1));
            return ReturnIgxlSheet(sheet, sheetType);
        }

        public IgxlSheet CreateIgxlSheet(ExcelWorksheet sheet)
        {
            var sheetType = GetIgxlSheetType(GetCellText(sheet, 1, 1));
            return ReturnIgxlSheet(sheet, sheetType);
        }

        private IgxlSheet ReturnIgxlSheet(Worksheet sheet, SheetType sheetType)
        {
            switch (sheetType)
            {
                case SheetType.DTFlowtableSheet:
                    return new ReadFlowSheet().GetSheet(sheet);
                case SheetType.DTTestInstancesSheet:
                    return new ReadInstanceSheet().GetSheet(sheet);
                case SheetType.DTLevelSheet:
                    return new ReadLevelSheet().GetSheet(sheet);
                case SheetType.DTPatternSetSheet:
                    return new ReadPatSetSheet().GetSheet(sheet);
                case SheetType.DTTimesetBasicSheet:
                    return new ReadTimeSetSheet().GetSheet(sheet);
                case SheetType.DTChanMap:
                    return new ReadChanMapSheet().GetSheet(sheet);
                case SheetType.DTBintablesSheet:
                    return new ReadBinTableSheet().GetSheet(sheet);
                case SheetType.DTJobListSheet:
                    return new ReadJobListSheet().GetSheet(sheet);
                case SheetType.DTGlobalSpecSheet:
                    return new ReadGlobalSpecSheet().GetSheet(sheet);
                case SheetType.DTACSpecSheet:
                    return new ReadAcSpecSheet().GetSheet(sheet);
                case SheetType.DTDCSpecSheet:
                    return new ReadDcSpecSheet().GetSheet(sheet);
                case SheetType.DTPinMap:
                    return new ReadPinMapSheet().GetSheet(sheet);
                case SheetType.DTUnknown:
                    break;
            }
            return null;
        }

        private IgxlSheet ReturnIgxlSheet(ExcelWorksheet sheet, SheetType sheetType)
        {
            switch (sheetType)
            {
                case SheetType.DTPortMapSheet:
                    return GetPortMapSheet(sheet);
                case SheetType.DTFlowtableSheet:
                    return new ReadFlowSheet().GetSheet(sheet);
                case SheetType.DTTestInstancesSheet:
                    return new ReadInstanceSheet().GetSheet(sheet);
                case SheetType.DTLevelSheet:
                    return new ReadLevelSheet().GetSheet(sheet);
                case SheetType.DTPatternSetSheet:
                    return new ReadPatSetSheet().GetSheet(sheet);
                case SheetType.DTTimesetBasicSheet:
                    return new ReadTimeSetSheet().GetSheet(sheet);
                case SheetType.DTChanMap:
                    return new ReadChanMapSheet().GetSheet(sheet);
                case SheetType.DTBintablesSheet:
                    return new ReadBinTableSheet().GetSheet(sheet);
                case SheetType.DTJobListSheet:
                    return new ReadJobListSheet().GetSheet(sheet);
                case SheetType.DTGlobalSpecSheet:
                    return new ReadGlobalSpecSheet().GetSheet(sheet);
                case SheetType.DTACSpecSheet:
                    return new ReadAcSpecSheet().GetSheet(sheet);
                case SheetType.DTDCSpecSheet:
                    return new ReadDcSpecSheet().GetSheet(sheet);
                case SheetType.DTPinMap:
                    return new ReadPinMapSheet().GetSheet(sheet);
                case SheetType.DTUnknown:
                    break;
            }
            return null;
        }

        public IgxlSheet CreateIgxlSheet(string file)
        {
            var firstLine = ReadFirstLine(file);
            var sheetType = GetIgxlSheetType(firstLine);
            var sheet = ConvertTxtToExcelSheet(file);
            return ReturnIgxlSheet(sheet, sheetType);
        }

        public ExcelWorksheet ConvertTxtToExcelSheet(string fileName)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            string sheetName = Path.GetFileNameWithoutExtension(fileName);
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            int index = 0;
            using (StreamReader sr = new StreamReader(fileName))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    index++;
                    if (line != null)
                    {
                        string[] arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        int cnt = 0;
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

        #region read Excel
        public List<IgxlSheet> GetIgxlSheets(ExcelWorkbook workbook, SheetType type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();
            foreach (var sheet in workbook.Worksheets)
            {
                if (GetIgxlSheetType(GetCellText(sheet, 1, 1)) == type)
                    igxlSheets.Add(CreateIgxlSheet(sheet));
            }
            return igxlSheets;
        }


        public List<IgxlSheet> GetIgxlSheets(Workbook workbook, SheetType type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //if (sheet.Cells[1, 1].Value2 != null &&
                //    sheet.Cells[1, 1].Value2.ToString().StartsWith("This sheet has not been populated.", StringComparison.CurrentCultureIgnoreCase))
                //    sheet.Select();
                if (GetIgxlSheetType(GetCellText(sheet, 1, 1)) == type)
                    igxlSheets.Add(CreateIgxlSheet(sheet));
            }
            return igxlSheets;
        }
        #endregion

        #region read *.txt
        public List<IgxlSheet> GetIgxlSheets(string path, SheetType type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();
            var files = Directory.GetFiles(path);
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file);
                if (extension == null || !extension.Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                var firstLine = ReadFirstLine(file);
                var sheetType = GetIgxlSheetType(firstLine);
                var version = GetIgxlSheetVersion(firstLine);
                if (sheetType == type)
                    igxlSheets.Add(CreateIgxlSheet(file));
            }
            return igxlSheets;
        }

        public List<IgxlSheet> GetIgxlSheets(List<string> sheetPaths, SheetType type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();

            foreach (var file in sheetPaths)
            {
                if (!File.Exists(file))
                    continue;

                var extension = Path.GetExtension(file);
                if (extension == null || !extension.Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                var firstLine = ReadFirstLine(file);
                var sheetType = GetIgxlSheetType(firstLine);
                var version = GetIgxlSheetVersion(firstLine);
                if (sheetType == type)
                    igxlSheets.Add(CreateIgxlSheet(file));
            }
            return igxlSheets;
        }

        public Dictionary<string, SheetType> GetSheetTypeDic(string path)
        {
            Dictionary<string, SheetType> dic = new Dictionary<string, SheetType>();
            var files = Directory.GetFiles(path, "*.txt", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file);
                if (extension == null || !extension.Equals(".txt", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                var firstLine = ReadFirstLine(file);
                var type = GetIgxlSheetType(firstLine);
                var version = GetIgxlSheetVersion(firstLine);
                if (!dic.ContainsKey(file))
                    dic.Add(file, type);
            }
            return dic;
        }

        public Dictionary<string, SheetType> GetSheetTypeDic(Workbook workbook)
        {
            Dictionary<string, SheetType> dic = new Dictionary<string, SheetType>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var value = sheet.Cells[1, 1].Value2;
                if (value != null)
                {
                    var type = GetIgxlSheetType(value.ToString());
                    if (!dic.ContainsKey(sheet.Name))
                        dic.Add(sheet.Name, type);
                }
            }
            return dic;
        }

        public bool IsSheetType(SheetType type, string file)
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

        public SheetType GetIgxlSheetType(string text)
        {
            if (Regex.IsMatch(text, SheetType.DTFlowtableSheet.ToString()))
                return SheetType.DTFlowtableSheet;
            if (Regex.IsMatch(text, SheetType.DTTestInstancesSheet.ToString()))
                return SheetType.DTTestInstancesSheet;
            if (Regex.IsMatch(text, SheetType.DTDCSpecSheet.ToString()))
                return SheetType.DTDCSpecSheet;
            if (Regex.IsMatch(text, SheetType.DTACSpecSheet.ToString()))
                return SheetType.DTACSpecSheet;
            if (Regex.IsMatch(text, SheetType.DTLevelSheet.ToString()))
                return SheetType.DTLevelSheet;
            if (Regex.IsMatch(text, SheetType.DTGlobalSpecSheet.ToString()))
                return SheetType.DTGlobalSpecSheet;
            if (Regex.IsMatch(text, SheetType.DTTimesetBasicSheet.ToString()))
                return SheetType.DTTimesetBasicSheet;
            if (Regex.IsMatch(text, SheetType.DTBintablesSheet.ToString()))
                return SheetType.DTBintablesSheet;
            if (Regex.IsMatch(text, SheetType.DTChanMap.ToString()))
                return SheetType.DTChanMap;
            if (Regex.IsMatch(text, SheetType.DTCharacterizationSheet.ToString()))
                return SheetType.DTCharacterizationSheet;
            if (Regex.IsMatch(text, SheetType.DTJobListSheet.ToString()))
                return SheetType.DTJobListSheet;
            if (Regex.IsMatch(text, SheetType.DTPatternSetSheet.ToString()))
                return SheetType.DTPatternSetSheet;
            if (Regex.IsMatch(text, SheetType.DTPatternSubroutineSheet.ToString()))
                return SheetType.DTPatternSubroutineSheet;
            if (Regex.IsMatch(text, SheetType.DTPinMap.ToString()))
                return SheetType.DTPinMap;
            if (Regex.IsMatch(text, SheetType.DTPortMapSheet.ToString()))
                return SheetType.DTPortMapSheet;

            return SheetType.DTUnknown;
        }

        private string ReadFirstLine(string file)
        {
            using (var sw = new StreamReader(file))
            {
                return sw.ReadLine() ?? "";
            }
        }
        #endregion

        public List<string> GetSheetsByType(Workbook workbook, SheetType type)
        {
            var dic = GetSheetTypeDic(workbook);
            return dic.ToList().Where(x => x.Value == type).Select(x => x.Key).ToList();
        }

        public List<string> GetSheetsByType(ExcelWorkbook workbook, SheetType type)
        {
            var sheets = workbook.Worksheets.Where(x => GetIgxlSheetType(GetCellText(x, 1, 1)) == type).Select(x => x.Name).ToList();
            return sheets;
        }

        public List<string> GetSheetsByType(string path, SheetType type)
        {
            var dic = GetSheetTypeDic(path);
            return dic.ToList().Where(x => x.Value == type).Select(x => x.Key).ToList();
        }

        public PortMapSheet GetPortMapSheet(ExcelWorksheet worksheet)
        {
            var firstLine = GetCellText(worksheet, 1, 1);
            var version = GetIgxlSheetVersion(firstLine);
            const string key = IgxlSheetNameList.PortMap;
            SheetObjMap sheetObjMap = null;
            if (_igxlConfigDic.ContainsKey(key))
                if (_igxlConfigDic[key].ContainsKey(version))
                    sheetObjMap = _igxlConfigDic[key][version];
            var readPortMapSheet = new ReadPortMapSheet(sheetObjMap);
            return readPortMapSheet.GetSheet(worksheet);
        }
    }
}