using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CommonReaderLib
{
    public class PmicIdsRow : MyRow
    {
        public string SourceSheetName { set; get; }
        public string Col1 { set; get; }
        public string Col2 { set; get; }
        public string Col3 { set; get; }

        public PmicIdsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }
    }

    public class PmicIdsSheet : MySheet
    {
        public List<PmicIdsRow> Rows { get; set; }

        public int IndexCol1 = -1;
        public int IndexCol2 = -1;
        public int IndexCol3 = -1;

        public PmicIdsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicIdsRow>();
        }
    }

    public class PmicIdsReader : MySheetReader
    {
        private string _sheetName;
        private PmicIdsSheet _pmicIdsSheet;

        private const string ConHeaderCol1 = "Col1";
        private const string ConHeaderCol2 = "Col2";
        private const string ConHeaderCol3 = "Col3";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _indexCol1 = -1;
        private int _indexCol2 = -1;
        private int _indexCol3 = -1;

        public PmicIdsSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null) return null;

            ExcelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _pmicIdsSheet = new PmicIdsSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _pmicIdsSheet = ReadSheet();

            return _pmicIdsSheet;
        }

        private PmicIdsSheet ReadSheet()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PmicIdsRow row = new PmicIdsRow(_sheetName);
                row.RowNum = i;
                if (_indexCol1 != -1)
                    row.Col1 = ExcelWorksheet.GetMergedCellValue(i, _indexCol1).Trim();
                if (_indexCol2 != -1)
                    row.Col2 = ExcelWorksheet.GetMergedCellValue(i, _indexCol2).Trim();
                if (_indexCol3 != -1)
                    row.Col3 = ExcelWorksheet.GetMergedCellValue(i, _indexCol3).Trim();
                _pmicIdsSheet.Rows.Add(row);
            }
            _pmicIdsSheet.IndexCol1 = _indexCol1;
            _pmicIdsSheet.IndexCol2 = _indexCol2;
            _pmicIdsSheet.IndexCol3 = _indexCol3;
            return _pmicIdsSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = ExcelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderCol1, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCol1 = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCol2, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCol2 = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCol3, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCol3 = i;
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().Equals(ConHeaderCol1, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _indexCol1 = -1;
            _indexCol2 = -1;
            _indexCol3 = -1;
        }
    }
}