using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace PmicAutomation.Utility.Relay.Input
{
    public class AdgMatrixRow
    {
        #region Constructor

        public AdgMatrixRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            AdgMatrix = "";
        }

        #endregion

        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string AdgMatrix { set; get; }

        #endregion
    }

    public class AdgMatrixSheet
    {
        #region Constructor

        public AdgMatrixSheet(string name)
        {
            Name = name;
            Rows = new List<AdgMatrixRow>();
            HeaderIndexDic = new Dictionary<string, int>();
        }

        #endregion

        #region Properity

        public string Name { get; set; }
        public List<AdgMatrixRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndexDic { get; set; }

        #endregion
    }

    public class AdgMatrixReader
    {
        private const string HeaderAdgMatrix = "AdgMatrix";

        private readonly Dictionary<string, bool> _headerDic = new Dictionary<string, bool> {{"AdgMatrix", true}};

        private int _adgMatrixIndex = -1;
        private AdgMatrixSheet _adgMatrixSheet;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private string _name;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public AdgMatrixSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _adgMatrixSheet = new AdgMatrixSheet(_name);

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

            _adgMatrixSheet = ReadSheetData();

            return _adgMatrixSheet;
        }

        private AdgMatrixSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                AdgMatrixRow row = new AdgMatrixRow(_name) {RowNum = i};
                if (_adgMatrixIndex != -1)
                {
                    row.AdgMatrix = _excelWorksheet.GetMergeCellValue(i, _adgMatrixIndex).Trim();
                }

                _adgMatrixSheet.Rows.Add(row);
            }

            return _adgMatrixSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderAdgMatrix, StringComparison.OrdinalIgnoreCase))
                {
                    _adgMatrixIndex = i;
                    _adgMatrixSheet.HeaderIndexDic.Add(HeaderAdgMatrix, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _adgMatrixSheet.HeaderIndexDic)
            {
                if (header.Value == -1 && _headerDic.ContainsKey(header.Key) && _headerDic[header.Key])
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
                    .Equals(HeaderAdgMatrix, StringComparison.OrdinalIgnoreCase))
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
            _adgMatrixIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> mappingDictionaries = new List<Dictionary<string, string>>();
            foreach (AdgMatrixRow row in _adgMatrixSheet.Rows)
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string> {{"AdgMatrix", row.AdgMatrix}};
                mappingDictionaries.Add(dictionary);
            }

            return mappingDictionaries;
        }
    }
}