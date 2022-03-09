using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;

namespace PmicAutogen.Config.ProjectConfig
{
    public class ProjectConfigSettingReader
    {
        private const string HeaderName = "Name";
        private const string HeaderGroupName = "GroupName";
        private const string HeaderDevice = "TabGroup";
        private const string HeaderDescription = "Description";
        private const string HeaderType = "Type";
        private const string HeaderMaxlength = "MaxLength";
        private const string HeaderDefault = "Default";
        private const string HeaderOptionValue1 = "OptionValue1";
        private const string HeaderOptionValue2 = "OptionValue2";
        private const string HeaderOptionValue3 = "OptionValue3";
        private const string HeaderOptionValue4 = "OptionValue4";
        private const string HeaderOptionValue5 = "OptionValue5";
        private const string HeaderOptionValue6 = "OptionValue6";
        private const string HeaderOptionValue7 = "OptionValue7";
        private const string HeaderOptionValue8 = "OptionValue8";
        private const string HeaderOptionValue9 = "OptionValue9";
        private const string HeaderOptionValue10 = "OptionValue10";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"Name", true},
            {"GroupName", true},
            {"Description", true},
            {"Type", true},
            {"MaxLength", true},
            {"Default", true},
            {"OptionValue1", true},
            {"OptionValue2", true},
            {"OptionValue3", true},
            {"OptionValue4", true},
            {"OptionValue5", true},
            {"OptionValue6", true},
            {"OptionValue7", true},
            {"OptionValue8", true},
            {"OptionValue9", true},
            {"OptionValue10", true}
        };

        private int _defaultIndex = -1;
        private int _descriptionIndex = -1;
        private int _deviceIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _groupNameIndex = -1;
        private int _maxlengthIndex = -1;
        private int _nameIndex = -1;
        private int _optionValue10Index = -1;
        private int _optionValue1Index = -1;
        private int _optionValue2Index = -1;
        private int _optionValue3Index = -1;
        private int _optionValue4Index = -1;
        private int _optionValue5Index = -1;
        private int _optionValue6Index = -1;
        private int _optionValue7Index = -1;
        private int _optionValue8Index = -1;
        private int _optionValue9Index = -1;
        private ProjectConfigSettingSheet _projectConfigSheet;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public ProjectConfigSettingSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _projectConfigSheet = new ProjectConfigSettingSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _projectConfigSheet = ReadSheetData();

            return _projectConfigSheet;
        }

        private ProjectConfigSettingSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new ProjectConfigSettingRow(_sheetName);
                row.RowNum = i;
                if (_nameIndex != -1)
                    row.Name = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _nameIndex).Trim();
                if (_groupNameIndex != -1)
                    row.GroupName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _groupNameIndex).Trim();
                if (_deviceIndex != -1)
                    row.TabGroup = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _deviceIndex).Trim();
                if (_descriptionIndex != -1)
                    row.Description = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _descriptionIndex).Trim();
                if (_typeIndex != -1)
                    row.Type = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
                if (_maxlengthIndex != -1)
                {
                    int value;
                    int.TryParse(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _maxlengthIndex).Trim(),
                        out value);
                    row.MaxLength = value;
                }

                if (_defaultIndex != -1)
                    row.Default = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _defaultIndex).Trim();
                if (_optionValue1Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue1Index)
                        .Trim());
                if (_optionValue2Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue2Index)
                        .Trim());
                if (_optionValue3Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue3Index)
                        .Trim());
                if (_optionValue4Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue4Index)
                        .Trim());
                if (_optionValue5Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue5Index)
                        .Trim());
                if (_optionValue6Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue6Index)
                        .Trim());
                if (_optionValue7Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue7Index)
                        .Trim());
                if (_optionValue8Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue8Index)
                        .Trim());
                if (_optionValue9Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue9Index)
                        .Trim());
                if (_optionValue10Index != -1)
                    row.OptionValue.Add(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionValue10Index)
                        .Trim());
                _projectConfigSheet.Rows.Add(row);
            }

            return _projectConfigSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _nameIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    _groupNameIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderGroupName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderDevice, StringComparison.OrdinalIgnoreCase))
                {
                    _deviceIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderDevice, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderDescription, StringComparison.OrdinalIgnoreCase))
                {
                    _descriptionIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderDescription, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderType, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderMaxlength, StringComparison.OrdinalIgnoreCase))
                {
                    _maxlengthIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderMaxlength, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderDefault, StringComparison.OrdinalIgnoreCase))
                {
                    _defaultIndex = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderDefault, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue1, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue1Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue1, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue2, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue2Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue2, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue3, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue3Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue3, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue4, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue4Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue4, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue5, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue5Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue5, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue6, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue6Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue6, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue7, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue7Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue7, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue8, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue8Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue8, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue9, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue9Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue9, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptionValue10, StringComparison.OrdinalIgnoreCase))
                {
                    _optionValue10Index = i;
                    _projectConfigSheet.HeaderIndex.Add(HeaderOptionValue10, i);
                }
            }

            foreach (var header in _projectConfigSheet.HeaderIndex)
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                    return false;

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
            for (var j = 1; j <= colNum; j++)
                if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim()
                    .Equals(HeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _startRowNumber = i;
                    return true;
                }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _nameIndex = -1;
            _groupNameIndex = -1;
            _deviceIndex = -1;
            _descriptionIndex = -1;
            _typeIndex = -1;
            _maxlengthIndex = -1;
            _defaultIndex = -1;
            _optionValue1Index = -1;
            _optionValue2Index = -1;
            _optionValue3Index = -1;
            _optionValue4Index = -1;
            _optionValue5Index = -1;
            _optionValue6Index = -1;
            _optionValue7Index = -1;
            _optionValue8Index = -1;
            _optionValue9Index = -1;
            _optionValue10Index = -1;
        }
    }

    public class ProjectConfigSettingRow
    {
        #region Constructor

        public ProjectConfigSettingRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Name = "";
            GroupName = "";
            TabGroup = "";
            Description = "";
            Type = "";
            Default = "";
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string Name { set; get; }
        public string GroupName { set; get; }
        public string TabGroup { set; get; }
        public string Description { set; get; }
        public string Type { set; get; }
        public int MaxLength { set; get; }
        public string Default { set; get; }
        public List<string> OptionValue = new List<string>();

        #endregion
    }

    public class ProjectConfigSettingSheet
    {
        #region Constructor

        public ProjectConfigSettingSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<ProjectConfigSettingRow>();
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<ProjectConfigSettingRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndex { get; set; } = new Dictionary<string, int>();

        #endregion
    }
}