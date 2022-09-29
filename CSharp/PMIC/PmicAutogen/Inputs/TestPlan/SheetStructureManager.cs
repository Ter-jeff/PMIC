using CommonLib.Extension;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using PmicAutogen.Config;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan
{
    public enum EnumColumn
    {
        None,
        Pin,
        PinMapType,
        ChannelMapType,
        Decimal,
        Int,
        DctestContinuityCategory,
        Formula,
        Unit
    }

    public class SheetStructureManager
    {
        private static readonly List<string> PinMapType = new List<string>
        {
            "I/O",
            "Input",
            "Output",
            "Analog",
            "Power",
            "Gnd",
            "Utility",
            "Voltage",
            "Current",
            "Unknown"
        };

        private static readonly List<string> ChannelMapType = new List<string>
        {
            "DCDiffMeter",
            "DCTime",
            "DCTimeHP",
            "DCTimeTrig",
            "DCVI",
            "DCVIMerged",
            "DCVS",
            "DCVSMerged2",
            "DCVSMerged4",
            "DCVSMerged6",
            "DCVSMerged8",
            "Gnd",
            "I/O",
            "N/C",
            "Utility"
        };

        private static readonly List<string> DctestContinuityCategoryList = new List<string> { "Continuity" };

        public static List<string> UnitList = new List<string>
            {"nV", "uV", "mV", "V", "nA", "uA", "mA", "A", "KHz", "MHz", "GHz", "Hz"};

        public static List<SheetConfig> SheetConfigs { get; set; }

        public static void Initialize()
        {
            if (SheetConfigs == null)
                SheetConfigs = GetSheetConfig();
        }

        public static List<SheetConfig> GetSheetConfig()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith(".Config.xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                    InputFiles.ConfigWorkbook = inputExcel.Workbook;
                    var workSheet = InputFiles.ConfigWorkbook.Worksheets[PmicConst.SheetConfig];
                    if (workSheet != null)
                    {
                        var projectConfigSettingReader = new SheetConfigReader();
                        var sheetConfigSheet = projectConfigSettingReader.ReadSheet(workSheet);
                        var sheetConfigs = new List<SheetConfig>();
                        foreach (var row in sheetConfigSheet.Rows)
                        {
                            var sheetConfig = new SheetConfig();
                            sheetConfig.SheetName = row.SheetName;
                            sheetConfig.FirstHeaderName = row.FirstHeaderName;
                            sheetConfig.HeaderName = row.HeaderName;
                            sheetConfig.Optional = row.Optional.Equals("True", StringComparison.CurrentCultureIgnoreCase);
                            sheetConfig.Type = GetColumnType(row);
                            sheetConfigs.Add(sheetConfig);
                        }
                        return sheetConfigs;
                    }
                }

            return null;
        }

        private static EnumColumn GetColumnType(SheetConfigRow row)
        {
            if (row.Type.Equals("Pin", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.Pin;
            else if (row.Type.Equals("PinMapType", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.PinMapType;
            else if (row.Type.Equals("ChannelMapType", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.ChannelMapType;
            else if (row.Type.Equals("Decimal", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.Decimal;
            else if (row.Type.Equals("Int", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.Int;
            else if (row.Type.Equals("DctestContinuityCategory", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.DctestContinuityCategory;
            else if (row.Type.Equals("Formula", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.Formula;
            else if (row.Type.Equals("Unit", StringComparison.CurrentCultureIgnoreCase))
                return EnumColumn.Unit;
            return EnumColumn.None;
        }

        #region Set validation

        public static void SetInitialColumns(Workbook workbook, Worksheet worksheet, string firstHeader,
            string targetHeader, EnumColumn columnType)
        {
            int headerRowNumber;
            var columnIndex = worksheet.GetColumnIndexByHeader(firstHeader, targetHeader, out headerRowNumber);
            if (columnIndex != -1)
            {
                worksheet.Unprotect();

                Validation validation = worksheet.Columns[columnIndex].Validation;

                if (columnType == EnumColumn.Pin)
                {
                    validation.Delete();
                    var sheet = workbook.GetSheet(PmicConst.IoPinMap);
                    if (sheet != null)
                        validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                            XlFormatConditionOperator.xlBetween, "=" + sheet.Name + "!" + GetRangeAddress(sheet));
                }
                else if (columnType == EnumColumn.PinMapType)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                        XlFormatConditionOperator.xlBetween, string.Join(",", PinMapType));
                }
                else if (columnType == EnumColumn.ChannelMapType)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                        XlFormatConditionOperator.xlBetween, string.Join(",", ChannelMapType));
                }
                else if (columnType == EnumColumn.Decimal)
                {
                    validation.Delete();
                    Range range = worksheet.Columns[columnIndex];
                    var letter = range.GetColumnLetter();
                    validation.Add(XlDVType.xlValidateCustom, XlDVAlertStyle.xlValidAlertStop,
                        XlFormatConditionOperator.xlBetween, "=IsNumber(" + letter + "1)");
                    validation.ErrorMessage = "It should be number !!!";
                }
                else if (columnType == EnumColumn.Int)
                {
                    validation.Delete();
                    Range range = worksheet.Columns[columnIndex];
                    var letter = range.GetColumnLetter();
                    validation.Add(XlDVType.xlValidateCustom, XlDVAlertStyle.xlValidAlertStop,
                        XlFormatConditionOperator.xlBetween, "=IsNumber(" + letter + "1)");
                    validation.ErrorMessage = "It should be number !!!";
                }
                else if (columnType == EnumColumn.DctestContinuityCategory)
                {
                    validation.Delete();
                    validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                        XlFormatConditionOperator.xlBetween, string.Join(",", DctestContinuityCategoryList));
                }
                else if (columnType == EnumColumn.Formula)
                {
                }
                else if (columnType == EnumColumn.Unit)
                {
                }

                worksheet.GetRange(1, columnIndex, headerRowNumber, columnIndex).Validation.Delete();
                // worksheet.Protect(Type.Missing, false, true, false, false, true, true, true, true, true, true, true,true, true, true, true);
            }
        }

        #endregion

        private static string GetRangeAddress(Worksheet worksheet)
        {
            int headerRowNumber;
            var columnIndex = worksheet.GetColumnIndexByHeader("Group Name", "Pin Name", out headerRowNumber);
            if (columnIndex != -1)
            {
                var finalRow = worksheet.Cells[headerRowNumber, columnIndex].End(XlDirection.xlDown).Row();
                return worksheet.Range[worksheet.Cells[headerRowNumber + 1, columnIndex],
                    worksheet.Cells[finalRow, columnIndex]].Address();
            }

            return "";
        }

        public static bool JudgeCell(EnumColumn columnType, string value, out string errorMessage)
        {
            errorMessage = "";

            if (string.IsNullOrEmpty(value)) return true;

            if (columnType == EnumColumn.Decimal)
            {
                double number;
                if (double.TryParse(value, out number))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not number !!!");
                return false;
            }

            if (columnType == EnumColumn.Int)
            {
                int number;
                if (int.TryParse(value, out number))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not integer !!!");
                return false;
            }

            if (columnType == EnumColumn.Formula)
            {
                if (IsFormula(value, ""))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value, "It is not formula !!!");
                return false;
            }

            if (columnType == EnumColumn.Unit)
            {
                if (IsUnit(value))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value,
                    "It is not value with unit !!!");
                return false;
            }

            if (columnType == EnumColumn.PinMapType)
            {
                if (PinMapType.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase)))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! => {1}", value, string.Join(",", PinMapType));
                return false;
            }

            if (columnType == EnumColumn.ChannelMapType)
            {
                if (ChannelMapType.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase)))
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value,
                    string.Join(",", ChannelMapType));
                return false;
            }

            if (columnType == EnumColumn.DctestContinuityCategory)
            {
                if (DctestContinuityCategoryList.Exists(x => x.Equals(value, StringComparison.CurrentCultureIgnoreCase))
                   )
                    return true;
                errorMessage = string.Format("\"{0}\" is not allowed !!! =>  {1}", value,
                    string.Join(",", ChannelMapType));
                return false;
            }

            return true;
        }

        private static bool IsUnit(string value)
        {
            try
            {
                value = value.Replace(" ", "");
                foreach (var unit in UnitList)
                    if (value.EndsWith(unit))
                    {
                        value = Regex.Replace(value, unit, "", RegexOptions.IgnoreCase);
                        break;
                    }

                double outValue;
                if (double.TryParse(value, out outValue))
                    return true;

                if (Regex.IsMatch(value, "[a-zA-Z]", RegexOptions.IgnoreCase))
                    return false;

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsFormula(string formula, string variable)
        {
            try
            {
                formula = !string.IsNullOrEmpty(variable)
                    ? formula.Replace(" ", "").TrimStart('=').Replace(variable, "1")
                    : formula.Replace(" ", "").TrimStart('=');

                if (Regex.IsMatch(formula, "[a-zA-Z]", RegexOptions.IgnoreCase))
                    return false;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

    public class SheetConfig
    {
        public SheetConfig()
        {
            SheetName = "";
            FirstHeaderName = "";
            HeaderName = "";
            Optional = false;
            Type = EnumColumn.None;
        }

        public string SheetName { get; set; }
        public string FirstHeaderName { get; set; }
        public string HeaderName { get; set; }
        public bool Optional { get; set; }
        public EnumColumn Type { get; set; }
    }
}