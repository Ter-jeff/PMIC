using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class IoPinGroupRow
    {
        #region Constructor

        public IoPinGroupRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            PinGroupName = "";
            PinNameContainedInPinGroup = "";
            Comments = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string PinGroupName { get; set; }
        public string PinNameContainedInPinGroup { get; set; }
        public string Comments { get; set; }

        #endregion
    }

    public class IoPinGroupSheet
    {
        #region Constructor

        public IoPinGroupSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<IoPinGroupRow>();
        }

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<IoPinGroupRow> Rows { get; }

        public int PinGroupNameIndex = -1;
        public int PinNameContainedInPinGroupIndex = -1;
        public int CommentsIndex = -1;

        #endregion
    }

    public class IoPinGroupReader
    {
        private const string HeaderPinGroupName = "Pin Group Name";
        private const string HeaderPinNameContainedInPinGroup = "Pin name Contained in Pin Group";
        private const string HeaderComments = "Comments";
        private int _commentsIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private IoPinGroupSheet _iOPinGroupSheet;
        private int _pinGroupNameIndex = -1;
        private int _pinNameContainedInPinGroupIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public IoPinGroupSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _iOPinGroupSheet = new IoPinGroupSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _iOPinGroupSheet = ReadSheetData();

            return _iOPinGroupSheet;
        }

        private IoPinGroupSheet ReadSheetData()
        {
            var ioPinGroupSheet = new IoPinGroupSheet(_sheetName);
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new IoPinGroupRow(_sheetName);
                row.RowNum = i;
                if (_pinGroupNameIndex != -1)
                    row.PinGroupName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _pinGroupNameIndex)
                        .Trim();
                if (_pinNameContainedInPinGroupIndex != -1)
                    row.PinNameContainedInPinGroup = EpplusOperation
                        .GetMergedCellValue(_excelWorksheet, i, _pinNameContainedInPinGroupIndex).Trim();
                if (_commentsIndex != -1)
                    row.Comments = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _commentsIndex).Trim();
                ioPinGroupSheet.Rows.Add(row);
            }

            ioPinGroupSheet.PinGroupNameIndex = _pinGroupNameIndex;
            ioPinGroupSheet.PinNameContainedInPinGroupIndex = _pinNameContainedInPinGroupIndex;
            ioPinGroupSheet.CommentsIndex = _commentsIndex;
            return ioPinGroupSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderPinGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    _pinGroupNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPinNameContainedInPinGroup, StringComparison.OrdinalIgnoreCase))
                {
                    _pinNameContainedInPinGroupIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderComments, StringComparison.OrdinalIgnoreCase)) _commentsIndex = i;
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
            for (var j = 1; j <= colNum; j++)
                if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim()
                    .Equals(HeaderPinGroupName, StringComparison.OrdinalIgnoreCase))
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
            _pinGroupNameIndex = -1;
            _pinNameContainedInPinGroupIndex = -1;
            _commentsIndex = -1;
        }
    }
}