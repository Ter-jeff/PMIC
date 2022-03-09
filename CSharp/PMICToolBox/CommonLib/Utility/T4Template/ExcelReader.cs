using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CommonLib.Utility
{
    public class PortDefineRow
    {
        #region Field
        public string SourceSheetName;
        public int RowNum;
        #endregion

        #region Properity
        public string ProtocolPortName { set; get; }
        public string Type { set; get; }
        public string Pin { set; get; }
        #endregion

        #region Constructor
        public PortDefineRow()
        {
        }

        public PortDefineRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }
        #endregion
    }

    public class PortDefineSheet
    {
        public string SheetName { get; set; }
        public List<PortDefineRow> Rows { set; get; }

        public int ProtocolPortNameIndex { set; get; }
        public int TypeIndex { set; get; }
        public int PinIndex { set; get; }

        #region Constructor
        public PortDefineSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = new List<PortDefineRow>();
        }
        #endregion
    }

    public class PortDefineReader
    {
	    private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
		private PortDefineSheet _portDefineSheet;

        private const string ConHeaderProtocolPortName = "Protocol Port Name";
        private const string ConHeaderType = "Type";
        private const string ConHeaderPin = "Pin";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _protocolPortNameIndex = -1;
        private int _typeIndex = -1;
        private int _pinIndex = -1;

        public PortDefineSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _portDefineSheet = new PortDefineSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            GetHeaderIndex();

            _portDefineSheet = ReadSheetData();

            return _portDefineSheet;
        }

        private PortDefineSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PortDefineRow(_sheetName);
                row.RowNum = i;
                if (_protocolPortNameIndex != -1)
                    row.ProtocolPortName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _protocolPortNameIndex).Trim();
                if (_typeIndex != -1)
                    row.Type = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _typeIndex).Trim();
                if (_pinIndex != -1)
                    row.Pin = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _pinIndex).Trim();
                _portDefineSheet.Rows.Add(row);
            }

            _portDefineSheet.ProtocolPortNameIndex = _protocolPortNameIndex;
            _portDefineSheet.TypeIndex = _typeIndex;
            _portDefineSheet.PinIndex = _pinIndex;

            return _portDefineSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string header = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (header.Equals(ConHeaderProtocolPortName, StringComparison.OrdinalIgnoreCase))
                {
                    _protocolPortNameIndex = i;
                    continue;
                }
                if (header.Equals(ConHeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    continue;
                }
                if (header.Equals(ConHeaderPin, StringComparison.OrdinalIgnoreCase))
				{
                    _pinIndex = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderProtocolPortName, StringComparison.OrdinalIgnoreCase))
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
            _protocolPortNameIndex = -1;
            _typeIndex = -1;
            _pinIndex = -1;
        }
    }
}