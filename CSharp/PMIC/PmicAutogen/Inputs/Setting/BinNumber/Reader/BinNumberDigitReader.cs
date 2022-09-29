using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.Setting.BinNumber.Reader
{
    public class BinNumberDigitReader
    {
        private const string Category = "Category";
        private const string Module = "Module";
        private const string SubModule = "SubModule";
        private const string Block = "Block";
        private const string Level = "Level";
        private const string Refer = @"See\s*(?<str>.*)";

        private List<string> _categoryList;
        private List<string> _moduleList;
        private List<string> _subModuleList;
        private List<string> _blockList;
        private List<string> _levelList;

        private void ReadModuleList(ExcelWorksheet moduleWorksheet)
        {
            int columnModule = -1;
            int columnSubModule = -1;
            int columnBlock = -1;
            int columnLevel = -1;
            int columnCategory = -1;
            _categoryList = new List<string>();
            _moduleList = new List<string>();
            _subModuleList = new List<string>();
            _blockList = new List<string>();
            _levelList = new List<string>();
            for (int i = 1; i <= moduleWorksheet.Dimension.End.Column; i++)
            {
                if (EpplusOperation.GetCellValue(moduleWorksheet, 1, i).Equals(Category, StringComparison.OrdinalIgnoreCase))
                {
                    columnCategory = i;
                }
                if (EpplusOperation.GetCellValue(moduleWorksheet, 1, i).Equals(Module, StringComparison.OrdinalIgnoreCase))
                {
                    columnModule = i;
                }
                else if (EpplusOperation.GetCellValue(moduleWorksheet, 1, i).Equals(SubModule, StringComparison.OrdinalIgnoreCase))
                {
                    columnSubModule = i;
                }
                else if (EpplusOperation.GetCellValue(moduleWorksheet, 1, i).Equals(Block, StringComparison.OrdinalIgnoreCase))
                {
                    columnBlock = i;
                }
                else if (EpplusOperation.GetCellValue(moduleWorksheet, 1, i).Equals(Level, StringComparison.OrdinalIgnoreCase))
                {
                    columnLevel = i;
                }
            }

            for (int i = 2; i <= moduleWorksheet.Dimension.End.Row; i++)
            {
                string category = ReadConfigCell(moduleWorksheet, i, columnCategory);
                string module = ReadConfigCell(moduleWorksheet, i, columnModule);
                string subModule = ReadConfigCell(moduleWorksheet, i, columnSubModule);
                string block = ReadConfigCell(moduleWorksheet, i, columnBlock);
                string level = ReadConfigCell(moduleWorksheet, i, columnLevel);
                if (!category.Equals(""))
                {
                    _categoryList.Add(category);
                }
                if (!module.Equals(""))
                {
                    _moduleList.Add(module);
                }
                if (!subModule.Equals(""))
                {
                    _subModuleList.Add(subModule);
                }
                if (!block.Equals(""))
                {
                    _blockList.Add(block);
                }
                if (!level.Equals(""))
                {
                    _levelList.Add(level);
                }
            }
        }

        public List<SoftBinDigitRow> ReadSheet(ExcelWorksheet worksheet)
        {
            List<SoftBinDigitRow> softBinData = new List<SoftBinDigitRow>();
            var moduleSheet = Input.ConfigWorkbook.Worksheets[PmicConst.ModuleList];
            if (moduleSheet != null)
            {
                ReadModuleList(moduleSheet);
                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    var range = worksheet.MergedCells[i, 1];
                    int start = i;
                    int end = range == null ? i : new ExcelAddress(range).End.Row;
                    SoftBinDigitRow digitalDef = ReadOneCat(worksheet, start, end, 4);
                    if (digitalDef != null)
                        softBinData.Add(digitalDef);
                }
            }
            return softBinData;
        }

        private SoftBinDigitRow ReadOneCat(ExcelWorksheet worksheet, int startRow, int endRow, int digitalNum)
        {
            SoftBinDigitRow softBinDigitalDef = new SoftBinDigitRow();
            string category = ReadConfigCell(worksheet, startRow, 1);
            softBinDigitalDef.Category = category;
            if (category.Equals("")) return null;
            softBinDigitalDef.CategoryType = GetBinNumKeyType(category);
            for (int i = 0; i < digitalNum; i++)
            {
                SoftBinDigit digital = new SoftBinDigit();
                for (int j = startRow; j <= endRow; j++)
                {
                    string keyword = EpplusOperation.GetMergedCellValue(worksheet, j, i * 2 + 2).Trim();
                    string number = EpplusOperation.GetCellValue(worksheet, j, i * 2 + 3).Trim();

                    if (Regex.IsMatch(keyword, Refer, RegexOptions.IgnoreCase))
                    {
                        var referSheetName = Regex.Match(keyword, Refer).Groups["str"].ToString();
                        if (!softBinDigitalDef.SoftBinDigits.Exists(p => referSheetName.Equals(p.ReferSheet)))
                        {
                            digital.ReferSheet = Regex.Match(keyword, Refer).Groups["str"].ToString();
                            var referSheet = Input.SettingWorkbook.Worksheets[digital.ReferSheet];
                            BinNumberMbistReader reader = new BinNumberMbistReader();
                            softBinDigitalDef.SoftBinMbists = reader.ReadSheet(referSheet);
                        }
                    }
                    if (!keyword.Equals("") && !digital.NumberInfos.Exists(p => p.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase)))
                    {
                        string keywordReplaceItemNameByConfigGroup = ProjectConfigSingleton.Instance().ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, keyword);
                        var type = GetBinNumKeyType(keywordReplaceItemNameByConfigGroup);
                        digital.AddNumDef(keyword, type, number);
                    }
                }
                softBinDigitalDef.SoftBinDigits.Add(digital);
            }

            return softBinDigitalDef;
        }

        private BinNumKeyType GetBinNumKeyType(string key)
        {
            if (_categoryList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return BinNumKeyType.Category;
            }
            if (_moduleList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return BinNumKeyType.Module;
            }
            if (_subModuleList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return BinNumKeyType.SubModule;
            }
            if (_blockList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return BinNumKeyType.Block;
            }
            if (_levelList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                return BinNumKeyType.Level;
            }
            return BinNumKeyType.NonType;
        }

        private string ReadConfigCell(ExcelWorksheet worksheet, int row, int column)
        {
            string result = EpplusOperation.GetCellValue(worksheet, row, column).Trim();
            result = ProjectConfigSingleton.Instance().ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, result);
            return result;
        }
    }
}