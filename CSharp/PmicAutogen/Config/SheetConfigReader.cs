using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;

namespace PmicAutogen.Config
{
    public class SheetConfigRow
    {
        #region Constructor

        public SheetConfigRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            SheetName = "";
            FirstHeaderName = "";
            HeaderName = "";
            Optional = "";
            Type = "";
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string SheetName { get; set; }
        public string FirstHeaderName { get; set; }
        public string HeaderName { get; set; }
        public string Optional { get; set; }
        public string Type { get; set; }

        #endregion
    }

    public class SheetConfigSheet
    {
        #region Constructor

        public SheetConfigSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<SheetConfigRow>();
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<SheetConfigRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndex { get; set; } = new Dictionary<string, int>();

        #endregion
    }

    public class SheetConfigReader
    {
        private const string HeaderSheetName = "SheetName";
        private const string HeaderFirstHeaderName = "FirstHeaderName";
        private const string HeaderHeaderName = "HeaderName";
        private const string HeaderOptional = "Optional";
        private const string HeaderType = "Type";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"SheetName", true},
            {"FirstHeaderName", true},
            {"HeaderName", true},
            {"Optional", true},
            {"Type", true}
        };

        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _firstHeaderNameIndex = -1;
        private int _headerNameIndex = -1;
        private int _optionalIndex = -1;
        private SheetConfigSheet _sheetConfigSheet;
        private string _sheetName;
        private int _sheetNameIndex = -1;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public SheetConfigSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _sheetConfigSheet = new SheetConfigSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _sheetConfigSheet = ReadSheetData();

            return _sheetConfigSheet;
        }

        private SheetConfigSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new SheetConfigRow(_sheetName);
                row.RowNum = i;
                if (_sheetNameIndex != -1)
                    row.SheetName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _sheetNameIndex).Trim();
                if (_firstHeaderNameIndex != -1)
                    row.FirstHeaderName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _firstHeaderNameIndex)
                        .Trim();
                if (_headerNameIndex != -1)
                    row.HeaderName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _headerNameIndex).Trim();
                if (_optionalIndex != -1)
                    row.Optional = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _optionalIndex).Trim();
                if (_typeIndex != -1)
                    row.Type = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
                _sheetConfigSheet.Rows.Add(row);
            }

            return _sheetConfigSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    _sheetNameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(HeaderSheetName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFirstHeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _firstHeaderNameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(HeaderFirstHeaderName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderHeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _headerNameIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(HeaderHeaderName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOptional, StringComparison.OrdinalIgnoreCase))
                {
                    _optionalIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(HeaderOptional, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    _sheetConfigSheet.HeaderIndex.Add(HeaderType, i);
                }
            }

            foreach (var header in _sheetConfigSheet.HeaderIndex)
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
                    .Equals(HeaderSheetName, StringComparison.OrdinalIgnoreCase))
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
            _sheetNameIndex = -1;
            _firstHeaderNameIndex = -1;
            _headerNameIndex = -1;
            _optionalIndex = -1;
            _typeIndex = -1;
        }
    }
}