using System.Xml.Serialization;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Teradyne.Oasis.IGData;
using Teradyne.Oasis.IGData.Utilities;
using PortMapSheet = IgxlData.IgxlSheets.PortMapSheet;

namespace IgxlData.IgxlReader
{
    public class IgxlSheetReader
    {
        private readonly Dictionary<string, Dictionary<string, SheetObjMap>> _igxlConfigDic = new Dictionary<string, Dictionary<string, SheetObjMap>>();

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

        private IGXL LoadConfig(Stream sr)
        {
            IGXL result;
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(IGXL));
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

        public IgxlSheet CreateIgxlSheet(ExcelWorksheet sheet)
        {
            var sheetType = GetIgxlSheetType(GetCellText(sheet, 1, 1));
            return ReturnIgxlSheet(sheet, sheetType);
        }

        private IgxlSheet ReturnIgxlSheet(ExcelWorksheet sheet, Sheet.SheetTypes sheetType)
        {
            switch (sheetType)
            {
                case Sheet.SheetTypes.DTPortMapSheet:
                    return GetPortMapSheet(sheet);
                case Sheet.SheetTypes.DTFlowtableSheet:
                    return new ReadFlowSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTTestInstancesSheet:
                    return new ReadInstanceSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTLevelSheet:
                    return new ReadLevelSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTPatternSetSheet:
                    return new ReadPatSetSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTTimesetBasicSheet:
                    return new ReadTimeSetSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTChanMap:
                    return new ReadChanMapSheet().ReadSheet(sheet);
                case Sheet.SheetTypes.DTJobListSheet:
                    return new ReadJobListSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTGlobalSpecSheet:
                    return new ReadGlobalSpecSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTACSpecSheet:
                    return new ReadAcSpecSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTDCSpecSheet:
                    return new ReadDcSpecSheet().GetSheet(sheet);
                case Sheet.SheetTypes.DTUnknown:
                    break;
            }
            return null;
        }

        public Sheet.SheetTypes GetIgxlSheetTypeByFile(string file)
        {
            var firstLine = ReadFirstLine(file);
            return  GetIgxlSheetType(firstLine);
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
        public List<IgxlSheet> GetIgxlSheets(ExcelWorkbook workbook, Sheet.SheetTypes type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();
            foreach (var sheet in workbook.Worksheets)
            {
                if (GetIgxlSheetType(GetCellText(sheet, 1, 1)) == type)
                    igxlSheets.Add(CreateIgxlSheet(sheet));
            }
            return igxlSheets;
        }

        public List<string> GetIgxlSheetList(ExcelWorkbook workbook, Sheet.SheetTypes type)
        {
            var sheets = workbook.Worksheets.Where(x => GetIgxlSheetType(GetCellText(x, 1, 1)) == type).Select(x => x.Name).ToList();
            return sheets;
        }
        #endregion

        #region read *.txt
        public List<IgxlSheet> GetIgxlSheets(string path, Sheet.SheetTypes type)
        {
            List<IgxlSheet> igxlSheets = new List<IgxlSheet>();
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

        public Dictionary<string, Sheet.SheetTypes> GetSheetTypeDic(string path)
        {
            Dictionary<string, Sheet.SheetTypes> dic = new Dictionary<string, Sheet.SheetTypes>();
            var files = Directory.GetFiles(path,"*.*",SearchOption.AllDirectories);//Modified by terry
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

        public List<string> GetSheetByType(string path, Sheet.SheetTypes type)
        {
            var dic = GetSheetTypeDic(path);
            return dic.ToList().Where(x => x.Value == type).Select(x => x.Key).ToList();
        }

        public bool IsSheetType(Sheet.SheetTypes type, string file)
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

        private Sheet.SheetTypes GetIgxlSheetType(string text)
        {
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTFlowtableSheet.ToString()))
                return Sheet.SheetTypes.DTFlowtableSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTTestInstancesSheet.ToString()))
                return Sheet.SheetTypes.DTTestInstancesSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTDCSpecSheet.ToString()))
                return Sheet.SheetTypes.DTDCSpecSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTACSpecSheet.ToString()))
                return Sheet.SheetTypes.DTACSpecSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTLevelSheet.ToString()))
                return Sheet.SheetTypes.DTLevelSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTGlobalSpecSheet.ToString()))
                return Sheet.SheetTypes.DTGlobalSpecSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTTimesetBasicSheet.ToString()))
                return Sheet.SheetTypes.DTTimesetBasicSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTBintablesSheet.ToString()))
                return Sheet.SheetTypes.DTBintablesSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTChanMap.ToString()))
                return Sheet.SheetTypes.DTChanMap;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTCharacterizationSheet.ToString()))
                return Sheet.SheetTypes.DTCharacterizationSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTJobListSheet.ToString()))
                return Sheet.SheetTypes.DTJobListSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTPatternSetSheet.ToString()))
                return Sheet.SheetTypes.DTPatternSetSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTPatternSubroutineSheet.ToString()))
                return Sheet.SheetTypes.DTPatternSubroutineSheet;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTPinMap.ToString()))
                return Sheet.SheetTypes.DTPinMap;
            if (Regex.IsMatch(text, Sheet.SheetTypes.DTPortMapSheet.ToString()))
                return Sheet.SheetTypes.DTPortMapSheet;

            return Sheet.SheetTypes.DTUnknown;
        }

        private string ReadFirstLine(string file)
        {
            using (var sw = new StreamReader(file))
            {
                return sw.ReadLine() ?? "";
            }
        }
        #endregion

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
    }
}