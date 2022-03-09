using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class OTPSetupRow
    {
        #region Property
        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string Variable { set; get; }
        public string Value { set; get; }
        public string Comment { set; get; }
        #endregion

        #region Constructor
        public OTPSetupRow()
        {
        }

        public OTPSetupRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }
        #endregion
    }

    public class OTPSetupSheet
    {
        #region Field
        internal int VaribleIndex;
        internal int ValueIndex;
        internal int CommentIndex;
        #endregion

        #region Property
        public string SheetName { get; set; }
        public List<OTPSetupRow> Rows { get; set; }
        #endregion

        #region Constructor
        public OTPSetupSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<OTPSetupRow>();
        }
        #endregion

        public string GetVaribleValue(string varibleName)
        {
            return Rows.Where(o => o.Variable.Equals(varibleName)).Select(o => o.Value).FirstOrDefault();
        }


        public Dictionary<string, string> GetJTAGPinNameMap()
        {
            string matchPattern = "^JTAG_([a-zA-Z]+)_Pin_Name$";
            Dictionary<string, string> map = new Dictionary<string, string>();
            Rows.ForEach(delegate (OTPSetupRow row)
            {
                var match = Regex.Match(row.Variable, matchPattern);
                if (match.Success
                    && !map.Keys.Contains(match.Groups[1].Value))
                {
                    map.Add(match.Groups[1].Value, row.Value);
                }
            });
            return map;
        }
    }

    public class OTPSetupReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private OTPSetupSheet _otpSetupSheet;

        private const string ConHeaderVarible = "Variable";
        private const string ConHeaderValue = "Value";
        private const string ConHeaderComment = "Comment";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _varibleIndex = -1;
        private int _valueIndex = -1;
        private int _commentIndex = -1;


        public OTPSetupSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _otpSetupSheet = new OTPSetupSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _otpSetupSheet = ReadSheetData();

            return _otpSetupSheet;
        }

        private OTPSetupSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                OTPSetupRow row = new OTPSetupRow(_sheetName);
                row.RowNum = i;
                if (_varibleIndex != -1)
                    row.Variable = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _varibleIndex).Trim();
                if (_valueIndex != -1)
                    row.Value = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _valueIndex).Trim();
                if (_commentIndex != -1)
                    row.Comment = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
                if (!string.IsNullOrEmpty(row.Variable))
                    _otpSetupSheet.Rows.Add(row);
            }

            _otpSetupSheet.VaribleIndex = _varibleIndex;
            _otpSetupSheet.ValueIndex = _valueIndex;
            _otpSetupSheet.CommentIndex = _commentIndex;

            return _otpSetupSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderVarible, StringComparison.OrdinalIgnoreCase))
                {
                    _varibleIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderValue, StringComparison.OrdinalIgnoreCase))
                {
                    _valueIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    continue;
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
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderVarible, StringComparison.OrdinalIgnoreCase))
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
            _varibleIndex = -1;
            _valueIndex = -1;
            _commentIndex = -1;
        }

    }
}