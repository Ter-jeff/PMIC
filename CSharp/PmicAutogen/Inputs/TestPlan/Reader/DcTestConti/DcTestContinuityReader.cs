using System;
using System.Text.RegularExpressions;
using AutomationCommon.Utility;
using OfficeOpenXml;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTestConti
{
    public class DcTestContinuityReader
    {
        #region Constructor

        public DcTestContinuityReader()
        {
            _startColNum = 1;
            _startRowNum = 1;
            _endColNum = 1;
            _endRowNum = 1;
            _limitValueNum = -1;
            _categoryIndex = -1;
            _pinGroupIndex = -1;
            _timeSetIndex = -1;
            _conditionIndex = -1;
            _limitIndex = -1;
        }

        #endregion

        #region public Function

        public DcTestContinuitySheet ReadSheet(ExcelWorksheet sheet)
        {
            try
            {
                _sheetName = sheet.Name;

                _sheet = sheet;
                _dcTestContiSheet = new DcTestContinuitySheet(_sheetName);

                GetDimension();

                ReadHeader();

                _dcTestContiSheet = ReadData();

                return _dcTestContiSheet;
            }
            catch (Exception)
            {
                throw new Exception("Meet an error in reading DcTestContiSheet");
            }
        }

        #endregion

        #region Get dimension

        private void GetDimension()
        {
            for (var i = 1; i < MaxSearchRow; i++)
            for (var j = 1; j < MaxSearchCol; j++)
            {
                var startHeader = GetCellValue(i, j);
                if (!startHeader.Equals("") && startHeader == DcTestContiRow.ConHeaderCategory)
                {
                    _startColNum = j;
                    _startRowNum = i;
                    break;
                }
            }

            _endColNum = _sheet.Dimension.End.Column;
            _endRowNum = _sheet.Dimension.End.Row;
        }

        #endregion

        #region Read Header

        private void ReadHeader()
        {
            for (var i = _startColNum; i <= _endColNum; i++)
            {
                var header = EpplusOperation.GetCellValue(_sheet, _startRowNum, i);
                if (IsLiked(header, DcTestContiRow.ConHeaderCategory))
                    _categoryIndex = i;
                else if (IsLiked(header, DcTestContiRow.ConHeaderPinGroup))
                    _pinGroupIndex = i;
                else if (IsLiked(header, DcTestContiRow.ConHeaderTimeSet))
                    _timeSetIndex = i;
                else if (IsLiked(header, DcTestContiRow.ConHeaderCondition))
                    _conditionIndex = i;
                else if (IsLiked(header, DcTestContiRow.ConHeaderLimit)) _limitIndex = i;

                if (IsLiked(header, DcTestContiRow.ConHeaderLimit))
                    _limitValueNum = GetMergedHeaderColRange(_startRowNum, i);
            }
        }

        #endregion

        #region Read Data

        private DcTestContinuitySheet ReadData()
        {
            var dcTestContinuitySheet = new DcTestContinuitySheet(_sheetName);
            var mergeCellCnt = 2;
            var range = _sheet.MergedCells[_startRowNum, _categoryIndex];
            if (range != null)
            {
                var lAddress = new ExcelAddress(range);
                mergeCellCnt = lAddress.End.Row - lAddress.Start.Row + 1;
            }

            for (var i = _startRowNum + mergeCellCnt; i <= _endRowNum; i++)
            {
                var dcTestContiRow = new DcTestContiRow();
                dcTestContiRow.Category = EpplusOperation.GetMergedCellValue(_sheet, i, _categoryIndex).Trim();
                dcTestContiRow.PinGroup = GetCellValue(i, _pinGroupIndex).Replace(" ", "_");
                if (_timeSetIndex != -1)
                    dcTestContiRow.TimeSet = GetCellValue(i, _timeSetIndex).Trim();
                dcTestContiRow.Condition = EpplusOperation.GetMergedCellValue(_sheet, i, _conditionIndex);

                if (dcTestContiRow.Category == "" && dcTestContiRow.PinGroup == "" && dcTestContiRow.Condition == "")
                    continue;

                dcTestContiRow.RowNum = i;
                dcTestContiRow.ColumnIdx = _pinGroupIndex;
                for (var j = 0; j < _limitValueNum; j += 4)
                {
                    var dcTestContiSheetLimit = new DcTestContiSheetLimit();
                    var col = j + _limitIndex;
                    var row = _startRowNum + 1;

                    if (EpplusOperation.GetMergedCellValue(_sheet, row, col).Contains(DcTestContiRow.ConHeaderHiShort))
                        dcTestContiSheetLimit.ShortHiLimitValue = EpplusOperation.GetMergedCellValue(_sheet, i, col);
                    if (EpplusOperation.GetMergedCellValue(_sheet, row, col + 1)
                        .Contains(DcTestContiRow.ConHeaderLoShort))
                        dcTestContiSheetLimit.ShortLoLimitValue =
                            EpplusOperation.GetMergedCellValue(_sheet, i, col + 1);
                    if (EpplusOperation.GetMergedCellValue(_sheet, row, col + 2).Equals(DcTestContiRow.ConHeaderHiOpen))
                        dcTestContiSheetLimit.OpenHiLimitValue = EpplusOperation.GetMergedCellValue(_sheet, i, col + 2);
                    if (EpplusOperation.GetMergedCellValue(_sheet, row, col + 3).Equals(DcTestContiRow.ConHeaderLoOpen))
                        dcTestContiSheetLimit.OpenLoLimitValue = EpplusOperation.GetMergedCellValue(_sheet, i, col + 3);
                    dcTestContiRow.LimitsPmic.Add(dcTestContiSheetLimit);
                }

                dcTestContinuitySheet.AddRow(dcTestContiRow);
            }

            dcTestContinuitySheet.CategoryIndex = _categoryIndex;
            dcTestContinuitySheet.PinGroupIndex = _pinGroupIndex;
            dcTestContinuitySheet.TimeSetIndex = _timeSetIndex;
            dcTestContinuitySheet.ConditionIndex = _conditionIndex;
            dcTestContinuitySheet.LimitIndex = _limitIndex;

            return dcTestContinuitySheet;
        }

        #endregion

        #region Field

        private const int MaxSearchRow = 10;
        private const int MaxSearchCol = 10;

        private ExcelWorksheet _sheet;
        private DcTestContinuitySheet _dcTestContiSheet;
        private int _startColNum;
        private int _startRowNum;
        private int _endColNum;
        private int _endRowNum;

        private int _categoryIndex;
        private int _pinGroupIndex;
        private int _timeSetIndex;
        private int _conditionIndex;
        private int _limitIndex;
        private int _limitValueNum;
        private string _sheetName;

        #endregion

        #region Common function

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            var value = _sheet.Cells[rowNumber, columnNumber].Value;
            if (value != null) return value.ToString().Trim();
            return "";
        }

        protected int GetMergedHeaderColRange(int pIRow, int pIColumn)
        {
            var range = _sheet.MergedCells[pIRow, pIColumn];
            if (range != null)
            {
                var lAddress = new ExcelAddress(range);
                return lAddress.End.Column - lAddress.Start.Column + 1;
            }

            return 1;
        }

        private string FormatStringForCompare(string pString)
        {
            var lStrResult = pString.Trim();

            lStrResult = ReplaceDoubleBlank(lStrResult);

            lStrResult = lStrResult.Replace(" ", "_");

            lStrResult = lStrResult.ToUpper();

            return lStrResult;
        }

        public string ReplaceDoubleBlank(string pString)
        {
            var lStrResult = pString;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        private bool IsLiked(string pStrInput, string pStrPatten)
        {
            if (pStrPatten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@".+", StringComparison.Ordinal) >= 0)
                return Regex.IsMatch(FormatStringForCompare(pStrInput), FormatStringForCompare(pStrPatten));
            return FormatStringForCompare(pStrInput) == FormatStringForCompare(pStrPatten);
        }

        #endregion
    }
}