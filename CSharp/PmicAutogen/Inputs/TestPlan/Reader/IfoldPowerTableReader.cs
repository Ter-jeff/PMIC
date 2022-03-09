using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class IfoldPowerTableRow
    {
        #region Constructor

        public IfoldPowerTableRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            PinName = "";
            Current = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string PinName { get; set; }
        public string Current { get; set; }

        #endregion
    }

    public class IfoldPowerTableSheet
    {
        #region Constructor

        public IfoldPowerTableSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<IfoldPowerTableRow>();
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<IfoldPowerTableRow> Rows { get; set; }

        public Dictionary<string, int> HeaderIndex { get; set; } = new Dictionary<string, int>();

        #endregion
    }

    public class IfoldPowerTableReader
    {
        private const string HeaderPinName = "PinName";
        private const string HeaderCurrentA = "Current (A)";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
            {{"PinName", true}, {"Current (A)", true}};

        private int _currentAIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _pinNameIndex = -1;
        private IfoldPowerTableSheet _powerTableSheet;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public IfoldPowerTableSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _powerTableSheet = new IfoldPowerTableSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _powerTableSheet = ReadSheetData();

            return _powerTableSheet;
        }

        private IfoldPowerTableSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new IfoldPowerTableRow(_sheetName);
                row.RowNum = i;
                if (_pinNameIndex != -1)
                    row.PinName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _pinNameIndex).Trim();
                if (_currentAIndex != -1)
                    row.Current = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _currentAIndex).Trim();
                _powerTableSheet.Rows.Add(row);
            }

            return _powerTableSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderPinName, StringComparison.OrdinalIgnoreCase))
                {
                    _pinNameIndex = i;
                    _powerTableSheet.HeaderIndex.Add(HeaderPinName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCurrentA, StringComparison.OrdinalIgnoreCase))
                {
                    _currentAIndex = i;
                    _powerTableSheet.HeaderIndex.Add(HeaderCurrentA, i);
                }
            }

            foreach (var header in _powerTableSheet.HeaderIndex)
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
                    .Equals(HeaderPinName, StringComparison.OrdinalIgnoreCase))
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
            _pinNameIndex = -1;
            _currentAIndex = -1;
        }
    }
}