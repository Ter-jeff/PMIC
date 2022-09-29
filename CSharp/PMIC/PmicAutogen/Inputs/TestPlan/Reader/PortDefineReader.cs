using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class PortDefineRow
    {
        #region Constructor

        public PortDefineRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            ProtocolPortName = "";
            Type = "";
            Pin = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string ProtocolPortName { get; set; }
        public string Type { get; set; }
        public string Pin { get; set; }

        #endregion
    }

    public class PortDefineSheet
    {
        #region Constructor

        public PortDefineSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PortDefineRow>();
        }

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<PortDefineRow> Rows { get; set; }
        public int ProtocolPortNameIndex = -1;
        public int TypeIndex = -1;
        public int PinIndex = -1;

        public Dictionary<string, List<PortDefineRow>> GroupByPortName()
        {
            return Rows.GroupBy(x => x.ProtocolPortName).ToDictionary(o => o.Key, o => o.ToList());
        }

        //Note that: If all pins in this port name, it's fine.
        public List<PortDefineRow> GetEmptyPinRows()
        {
            var invalidRows = new List<PortDefineRow>();
            var portAndPinDic = GroupByPortName();
            foreach (var portAndPinItem in portAndPinDic)
            {
                var emptyPinRows = portAndPinItem.Value.FindAll(o => string.IsNullOrEmpty(o.Pin.Trim()));
                if (emptyPinRows.Count > 0 && emptyPinRows.Count < portAndPinItem.Value.Count)
                    invalidRows.AddRange(emptyPinRows);
            }

            return invalidRows;
        }

        public string GetFirstPin(string name)
        {
            var fisrt = Rows.Where(x => x.ProtocolPortName.Equals(name, StringComparison.CurrentCultureIgnoreCase));
            if (fisrt != null && fisrt.Count() > 0)
                return fisrt.First().Pin;
            return "";
        }

        #endregion
    }

    public class PortDefineReader
    {
        private const string HeaderProtocolPortName = "Protocol Port Name";
        private const string HeaderType = "Type";
        private const string HeaderPin = "Pin";
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _pinIndex = -1;
        private PortDefineSheet _portDefineSheet;
        private int _protocolPortNameIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public PortDefineSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _portDefineSheet = new PortDefineSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _portDefineSheet = ReadSheetData();

            return _portDefineSheet;
        }

        private PortDefineSheet ReadSheetData()
        {
            var portDefineSheet = new PortDefineSheet(_sheetName);
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PortDefineRow(_sheetName);
                row.RowNum = i;
                if (_protocolPortNameIndex != -1)
                    row.ProtocolPortName = _excelWorksheet.GetMergedCellValue(i, _protocolPortNameIndex).Trim();
                if (_typeIndex != -1)
                    row.Type = _excelWorksheet.GetMergedCellValue(i, _typeIndex).Trim();
                if (_pinIndex != -1)
                    row.Pin = _excelWorksheet.GetMergedCellValue(i, _pinIndex).Trim();
                portDefineSheet.Rows.Add(row);
            }

            portDefineSheet.ProtocolPortNameIndex = _protocolPortNameIndex;
            portDefineSheet.TypeIndex = _typeIndex;
            portDefineSheet.PinIndex = _pinIndex;

            return portDefineSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderProtocolPortName, StringComparison.OrdinalIgnoreCase))
                {
                    _protocolPortNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPin, StringComparison.OrdinalIgnoreCase)) _pinIndex = i;
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(HeaderProtocolPortName, StringComparison.OrdinalIgnoreCase))
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
            _protocolPortNameIndex = -1;
            _typeIndex = -1;
            _pinIndex = -1;
        }
    }
}