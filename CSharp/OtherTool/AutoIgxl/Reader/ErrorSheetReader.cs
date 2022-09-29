using CommonReaderLib;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace AutoIgxl.Reader
{
    public class ErrorSheetReader : MySheetReader
    {
        private const string ConHeaderSheetName = "Sheet Name";
        private const string ConHeaderCell = "Cell";
        private const string ConHeaderErrorCode = "Error Code";
        private const string ConHeaderErrorMessage = "Error Message";
        private int _indexCell = -1;
        private int _indexErrorCode = -1;
        private int _indexErrorMessage = -1;

        private int _indexSheetName = -1;

        public ErrorSheet ReadSheet(Worksheet worksheet)
        {
            var sheetName = worksheet.Name;

            var sheet = new ErrorSheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                sheet.AddDimensionError();
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                sheet.AddFirstHeaderError(ConHeaderSheetName);
                return null;
            }

            GetHeaderIndex();

            sheet = ReadSheet(sheetName);

            return sheet;
        }

        private ErrorSheet ReadSheet(string sheetName)
        {
            var ruleSheet = new ErrorSheet(sheetName);
            for (var i = StartRowNumber + 1; i <= EndRowNumber; i++)
            {
                var row = new ErrorSheetRow(sheetName);
                row.RowNum = i;
                if (_indexSheetName != -1)
                    row.SheetName = GetMergedCellValue(ExcelWorksheet, i, _indexSheetName).Trim();
                if (_indexCell != -1)
                    row.Cell = GetMergedCellValue(ExcelWorksheet, i, _indexCell).Trim();
                if (_indexErrorCode != -1)
                    row.ErrorCode = GetMergedCellValue(ExcelWorksheet, i, _indexErrorCode).Trim();
                if (_indexErrorMessage != -1)
                    row.ErrorMessage = GetMergedCellValue(ExcelWorksheet, i, _indexErrorMessage).Trim();
                if (!string.IsNullOrEmpty(row.ErrorMessage))
                    ruleSheet.Rows.Add(row);
            }

            ruleSheet.IndexSheetName = _indexSheetName;
            ruleSheet.IndexCell = _indexCell;
            ruleSheet.IndexErrorCode = _indexErrorCode;
            ruleSheet.IndexErrorMessage = _indexErrorMessage;

            return ruleSheet;
        }

        private void GetHeaderIndex()
        {
            for (var i = StartColNumber; i <= EndColNumber; i++)
            {
                var header = GetCellValue(ExcelWorksheet, StartRowNumber, i).Trim();
                if (header.Equals(ConHeaderSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    _indexSheetName = i;
                    continue;
                }

                if (header.Equals(ConHeaderCell, StringComparison.OrdinalIgnoreCase))
                {
                    _indexCell = i;
                    continue;
                }

                if (header.Equals(ConHeaderErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    _indexErrorCode = i;
                    continue;
                }

                if (header.Equals(ConHeaderErrorMessage, StringComparison.OrdinalIgnoreCase)) _indexErrorMessage = i;
            }
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            var colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (GetCellValue(ExcelWorksheet, i, j).Trim()
                        .Equals(ConHeaderSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRowNumber = i;
                        return true;
                    }

            return false;
        }
    }

    public class ErrorSheet : MySheet
    {
        public int IndexCell = -1;
        public int IndexErrorCode = -1;
        public int IndexErrorMessage = -1;

        public int IndexSheetName = -1;

        #region Constructor

        public ErrorSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<ErrorSheetRow>();
        }

        #endregion

        public List<ErrorSheetRow> Rows { set; get; }
    }

    public class ErrorSheetRow : MyRow
    {
        #region Constructor

        public ErrorSheetRow(string sheetName = "")
        {
            SheetName = sheetName;
        }

        #endregion

        public string Cell { set; get; }
        public string ErrorCode { set; get; }
        public string ErrorMessage { set; get; }
    }
}