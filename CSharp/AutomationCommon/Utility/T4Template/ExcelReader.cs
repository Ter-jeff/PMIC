using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace AutomationCommon.Utility.T4Template
{
    public class PmicIdsRow
    {
        #region Property
        public string SourceSheetName { set ; get ; }
        public int RowNum { get ; set ; }
        public string InstanceName { set ; get ; }
        public string MeaseurePin { set ; get ; }
        public string CPFTQASNV { set ; get ; }
        public string QACNV { set ; get ; }
        public string CHARNV { set ; get ; }
        public string CHARLV { set ; get ; }
        public string CHARHV { set ; get ; }
        public string CHARULV { set ; get ; }
        #endregion

        #region Constructor
        public PmicIdsRow()
        {
        }

        public PmicIdsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }
        #endregion
    }

    public class PmicIdsSheet
    {
        #region Field
        private readonly Dictionary<string, int> _headerIndex = new Dictionary<string, int>();
        #endregion

        #region Property
        public string SheetName { get; set; }
        public List<PmicIdsRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndex { get { return _headerIndex; } }
        #endregion

        #region Constructor
        public PmicIdsSheet(string sheetName)
        {
            SheetName = sheetName;
			Rows = new List< PmicIdsRow>();
        }
        #endregion
    }

    public class PmicIdsReader
    {
	    private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
		private PmicIdsSheet _pmicIdsSheet;

        private const string ConHeaderInstanceName = "Instance Name";
        private const string ConHeaderMeaseurePin = "Measeure Pin";
        private const string ConHeaderCPFTQASNV = "CP_FT_QAS_NV";
        private const string ConHeaderQACNV = "QAC_NV";
        private const string ConHeaderCHARNV = "CHAR_NV";
        private const string ConHeaderCHARLV = "CHAR_LV";
        private const string ConHeaderCHARHV = "CHAR_HV";
        private const string ConHeaderCHARULV = "CHAR_ULV";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _instanceNameIndex = -1;
        private int _measeurePinIndex = -1;
        private int _cPFTQASNVIndex = -1;
        private int _qACNVIndex = -1;
        private int _cHARNVIndex = -1;
        private int _cHARLVIndex = -1;
        private int _cHARHVIndex = -1;
        private int _cHARULVIndex = -1;
        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool> 
		{
            { "Instance Name", true }, { "Measeure Pin", true }, { "CP_FT_QAS_NV", true }, { "QAC_NV", true }, { "CHAR_NV", true }, { "CHAR_LV", true }, { "CHAR_HV", true }, { "CHAR_ULV", true }
		};

        public PmicIdsSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _pmicIdsSheet = new PmicIdsSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _pmicIdsSheet = ReadSheetData();

            return _pmicIdsSheet;
        }

        private PmicIdsSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PmicIdsRow row = new PmicIdsRow(_sheetName);
                row.RowNum = i;
                if (_instanceNameIndex != -1)
                    row.InstanceName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _instanceNameIndex).Trim();
                if (_measeurePinIndex != -1)
                    row.MeaseurePin = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _measeurePinIndex).Trim();
                if (_cPFTQASNVIndex != -1)
                    row.CPFTQASNV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cPFTQASNVIndex).Trim();
                if (_qACNVIndex != -1)
                    row.QACNV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _qACNVIndex).Trim();
                if (_cHARNVIndex != -1)
                    row.CHARNV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cHARNVIndex).Trim();
                if (_cHARLVIndex != -1)
                    row.CHARLV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cHARLVIndex).Trim();
                if (_cHARHVIndex != -1)
                    row.CHARHV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cHARHVIndex).Trim();
                if (_cHARULVIndex != -1)
                    row.CHARULV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cHARULVIndex).Trim();
                _pmicIdsSheet.Rows.Add(row);
            }
            return _pmicIdsSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
                {
                    _instanceNameIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderInstanceName, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderMeaseurePin, StringComparison.OrdinalIgnoreCase))
                {
                    _measeurePinIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderMeaseurePin, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCPFTQASNV, StringComparison.OrdinalIgnoreCase))
                {
                    _cPFTQASNVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderCPFTQASNV, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderQACNV, StringComparison.OrdinalIgnoreCase))
                {
                    _qACNVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderQACNV, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARNV, StringComparison.OrdinalIgnoreCase))
                {
                    _cHARNVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderCHARNV, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARLV, StringComparison.OrdinalIgnoreCase))
                {
                    _cHARLVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderCHARLV, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARHV, StringComparison.OrdinalIgnoreCase))
                {
                    _cHARHVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderCHARHV, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARULV, StringComparison.OrdinalIgnoreCase))
                {
                    _cHARULVIndex = i;
                    _pmicIdsSheet.HeaderIndex.Add(ConHeaderCHARULV, i);
                    continue;
                }
            }

            foreach (var header in _pmicIdsSheet.HeaderIndex)
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                    return false;

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
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
            _instanceNameIndex = -1;
            _measeurePinIndex = -1;
            _cPFTQASNVIndex = -1;
            _qACNVIndex = -1;
            _cHARNVIndex = -1;
            _cHARLVIndex = -1;
            _cHARHVIndex = -1;
            _cHARULVIndex = -1;
        }

    }
}