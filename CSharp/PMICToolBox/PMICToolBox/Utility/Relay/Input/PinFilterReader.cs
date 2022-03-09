using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace PmicAutomation.Utility.Relay.Input
{
    public class PinFilterRow
    {
        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string Field { set; get; }
        public string Equal { set; get; }
        public string NotEqual { set; get; }
        public string Contain { set; get; }
        public string NotContain { set; get; }
        public string Prefixed { set; get; }
        public string Suffixed { set; get; }
        public string NotPrefixed { set; get; }
        public string NotSuffixed { set; get; }

        #endregion

        #region Constructor

        public PinFilterRow()
        {
            Field = "";
            Equal = "";
            NotEqual = "";
            Contain = "";
            NotContain = "";
            Prefixed = "";
            Suffixed = "";
            NotPrefixed = "";
            NotSuffixed = "";
        }

        public PinFilterRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Equal = "";
            NotEqual = "";
            Field = "";
            Contain = "";
            NotContain = "";
            Prefixed = "";
            Suffixed = "";
            NotPrefixed = "";
            NotSuffixed = "";
        }

        #endregion
    }

    public class PinFilterSheet
    {
        #region Constructor

        public PinFilterSheet(string name)
        {
            Name = name;
            Rows = new List<PinFilterRow>();
        }

        #endregion

        #region Field

        #endregion

        #region Properity

        public string Name { get; set; }
        public List<PinFilterRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndexDic= new Dictionary<string, int>();

        #endregion
    }

    public class PinFilterReader
    {
        private const string HeaderField = "Field";
        private const string HeaderEqual = "Equal";
        private const string HeaderContain = "Contain";
        private const string HeaderStartWith = "Prefix";
        private const string HeaderEndWith = "Suffix";
        private const string HeaderNotEqual = "Not Equal";
        private const string HeaderNotContain = "Not Contain";
        private const string HeaderNotStartWith = "Not Prefix";
        private const string HeaderNotEndWith = "Not Suffix";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"Field", true},
            {"Equal", true},
            {"Not Equal", true},
            {"Contain", true},
            {"Not Contain", true},
            {"Prefix", true},
            {"Suffix", true},
            {"Not Prefix", true},
            {"Not Suffix", true}
        };

        private ExcelWorksheet _excelWorksheet;
        private PinFilterSheet _pinFilterSheet;

        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        private string _name;
        private int _fieldIndex = -1;
        private int _equalIndex = -1;
        private int _notEqualIndex = -1;
        private int _containIndex = -1;
        private int _notContainIndex = -1;
        private int _prefixedIndex = -1;
        private int _notPrefixedIndex = -1;
        private int _suffixedIndex = -1;
        private int _notSuffixedIndex = -1;


        public PinFilterSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _pinFilterSheet = new PinFilterSheet(_name);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _pinFilterSheet = ReadSheetData();

            return _pinFilterSheet;
        }

        private PinFilterSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PinFilterRow row = new PinFilterRow(_name) {RowNum = i};
                if (_fieldIndex != -1)
                {
                    row.Field = _excelWorksheet.GetMergeCellValue(i, _fieldIndex).Trim();
                }

                if (_equalIndex != -1)
                {
                    row.Equal = _excelWorksheet.GetMergeCellValue(i, _equalIndex).Trim();
                }

                if (_notEqualIndex != -1)
                {
                    row.NotEqual = _excelWorksheet.GetMergeCellValue(i, _notEqualIndex).Trim();
                }

                if (_containIndex != -1)
                {
                    row.Contain = _excelWorksheet.GetMergeCellValue(i, _containIndex).Trim();
                }

                if (_notContainIndex != -1)
                {
                    row.NotContain = _excelWorksheet.GetMergeCellValue(i, _notContainIndex).Trim();
                }

                if (_prefixedIndex != -1)
                {
                    row.Prefixed = _excelWorksheet.GetMergeCellValue(i, _prefixedIndex).Trim();
                }

                if (_suffixedIndex != -1)
                {
                    row.Suffixed = _excelWorksheet.GetMergeCellValue(i, _suffixedIndex).Trim();
                }

                if (_notPrefixedIndex != -1)
                {
                    row.NotPrefixed = _excelWorksheet.GetMergeCellValue(i, _notPrefixedIndex).Trim();
                }

                if (_notSuffixedIndex != -1)
                {
                    row.NotSuffixed = _excelWorksheet.GetMergeCellValue(i, _notSuffixedIndex).Trim();
                }

                if (!string.IsNullOrEmpty(row.Field))
                {
                    _pinFilterSheet.Rows.Add(row);
                }
            }

            return _pinFilterSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderField, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderField, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderEqual, StringComparison.OrdinalIgnoreCase))
                {
                    _equalIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderEqual, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNotEqual, StringComparison.OrdinalIgnoreCase))
                {
                    _notEqualIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderNotEqual, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderContain, StringComparison.OrdinalIgnoreCase))
                {
                    _containIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderContain, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNotContain, StringComparison.OrdinalIgnoreCase))
                {
                    _notContainIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderNotContain, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderStartWith, StringComparison.OrdinalIgnoreCase))
                {
                    _prefixedIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderStartWith, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderEndWith, StringComparison.OrdinalIgnoreCase))
                {
                    _suffixedIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderEndWith, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNotStartWith, StringComparison.OrdinalIgnoreCase))
                {
                    _notPrefixedIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderNotStartWith, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNotEndWith, StringComparison.OrdinalIgnoreCase))
                {
                    _notSuffixedIndex = i;
                    _pinFilterSheet.HeaderIndexDic.Add(HeaderNotEndWith, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _pinFilterSheet.HeaderIndexDic)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            for (int j = 1; j <= colNum; j++)
            {
                if (_excelWorksheet.GetMergeCellValue(i, j).Trim()
                    .Equals(HeaderField, StringComparison.OrdinalIgnoreCase))
                {
                    _startRowNumber = i;
                    return true;
                }
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
            _fieldIndex = -1;
            _equalIndex = -1;
            _notEqualIndex = -1;
            _containIndex = -1;
            _notContainIndex = -1;
            _prefixedIndex = -1;
            _suffixedIndex = -1;
            _notPrefixedIndex = -1;
            _notSuffixedIndex = -1;
        }
    }
}