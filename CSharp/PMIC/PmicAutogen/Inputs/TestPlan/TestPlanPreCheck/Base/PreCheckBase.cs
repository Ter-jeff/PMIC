//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore. 
//
//------------------------------------------------------------------------------ 

using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base
{
    public abstract class PreCheckBase
    {
        private const int MaxStartRowIndex = 10;
        private const int MaxStartColumnIndex = 10;
        protected readonly List<SheetConfig> SheetConfigs = new List<SheetConfig>();
        protected readonly string SheetName;
        private bool _isBlankSheet;

        protected string FirstHeader = "";

        // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add Start
        protected bool IgnoreBlankSheet = false;
        protected int StartColumn;
        protected int StartRow;
        public int StopColumn = -1;

        public int StopRow = -1;
        // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add End

        protected ExcelWorkbook Workbook;
        protected ExcelWorksheet _excelWorksheet;

        #region Constructor

        protected PreCheckBase(ExcelWorkbook excelWorkbook, string sheetName)
        {
            Workbook = excelWorkbook;
            SheetName = sheetName;

            if (SheetStructureManager.SheetConfigs.Exists(x =>
                    x.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase)))
                SheetConfigs = SheetStructureManager.SheetConfigs
                    .Where(x => x.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase)).ToList();
            else
                foreach (var sheetConfig in SheetStructureManager.SheetConfigs)
                    if (sheetConfig.SheetName.Contains("*"))
                        if (Regex.IsMatch(sheetName, "^" + sheetConfig.SheetName.TrimEnd('*'), RegexOptions.IgnoreCase))
                            SheetConfigs.Add(sheetConfig);
        }

        #endregion

        protected abstract bool CheckBusiness();

        #region Member Function

        #region public method

        public bool CheckMain()
        {
            var checkResult = CheckExist();
            if (checkResult == false)
                return false;

            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Chg Start
            if (IgnoreBlankSheet && _isBlankSheet) return true;
            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Chg End

            CheckFormat();

            CheckBusiness();

            return true;
        }

        protected bool IsLiked(string input, string pattern)
        {
            if (pattern.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pattern.IndexOf(@".+", StringComparison.Ordinal) >= 0)
                return Regex.IsMatch(FormatStringForCompare(input), FormatStringForCompare(pattern));
            return FormatStringForCompare(input) == FormatStringForCompare(pattern);
        }

        #endregion

        #region private method

        private bool CheckExist()
        {
            var checkResult = CheckSheetName();
            if (checkResult == false)
                return false;

            if (SheetConfigs == null) return false;

            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add Start
            if (IgnoreBlankSheet) _isBlankSheet = CheckIsBlankSheet();

            if (IgnoreBlankSheet && _isBlankSheet) return true;
            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add End

            checkResult = CheckFirstHeaderLocation();
            if (checkResult == false)
                return false;

            checkResult = CheckHeaders();
            if (checkResult == false)
                return false;

            return true;
        }

        private bool CheckHeaders()
        {
            var headers = SheetConfigs.Where(x => x.HeaderName != null && !x.Optional).Select(x => x.HeaderName)
                .ToList();
            var result = true;
            for (var i = 0; i <= headers.Count - 1; i++)
            {
                var sourceHead = headers[i];
                var oneResult = false;
                for (var j = StartColumn; j <= _excelWorksheet.Dimension.End.Column; j++)
                {
                    var value = _excelWorksheet.GetCellValue(StartRow, j);
                    if (IsLiked(value, sourceHead))
                    {
                        oneResult = true;
                        break;
                    }
                }

                if (oneResult == false)
                {
                    var errorMessage = "Must Exist Item :" + sourceHead + " Not Exist.";
                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, StartRow, 1,
                        errorMessage);
                    result = false;
                }
            }

            return result;
        }

        protected virtual void CheckFormat()
        {
            foreach (var sheetConfig in SheetConfigs)
            {
                if (sheetConfig.Type == EnumColumn.None)
                    continue;
                var columnIndex = GetColumnIndex(sheetConfig);
                if (columnIndex != -1)
                    for (var i = StartRow + 1; i <= _excelWorksheet.Dimension.End.Row; i++)
                    {
                        string errorMessage;
                        var value = _excelWorksheet.GetCellValue(i, columnIndex);
                        if (!SheetStructureManager.JudgeCell(sheetConfig.Type, value, out errorMessage))
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, i,
                                columnIndex, errorMessage);
                    }
            }
        }

        protected int GetColumnIndex(SheetConfig sheetConfig)
        {
            for (var j = StartColumn; j <= _excelWorksheet.Dimension.End.Column; j++)
            {
                var value = _excelWorksheet.GetCellValue(StartRow, j);
                if (IsLiked(value, sheetConfig.HeaderName)) return j;
            }

            return -1;
        }

        private bool CheckSheetName()
        {
            if (Workbook.Worksheets.Any(
                    o => string.Equals(o.Name, SheetName, StringComparison.CurrentCultureIgnoreCase)))
                _excelWorksheet = Workbook.Worksheets[SheetName];
            else
                //var errorMessage = "Sheet: " + SheetName + " Not Exist.";
                //EpplusErrorManager.AddError(EnumErrorType.Existential, ErrorLevel.Error, SheetName, 1, 1, errorMessage);
                return false;

            return true;
        }

        private bool CheckFirstHeaderLocation()
        {
            StopRow = _excelWorksheet.Dimension.End.Row;
            StopColumn = _excelWorksheet.Dimension.End.Column;
            FirstHeader = SheetConfigs.First().FirstHeaderName;
            if (FirstHeader != "")
            {
                for (var i = StartRow; i <= MaxStartRowIndex; i++)
                    for (var j = StartColumn; j <= MaxStartColumnIndex; j++)
                    {
                        var value = _excelWorksheet.GetCellValue(i, j);
                        if (IsLiked(value, FirstHeader))
                        {
                            StartRow = i;
                            StartColumn = j;
                            return true;
                        }
                    }

                var errorMessage = "First Header Key Don't Exist In Sheet. Key:" + FirstHeader;
                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, 1, 1,
                    errorMessage);
                return false;
            }

            return true;
        }

        private string FormatStringForCompare(string value)
        {
            var result = value.Trim();

            result = ReplaceDoubleBlank(result);

            result = result.ToUpper();

            return result;
        }

        private string ReplaceDoubleBlank(string value)
        {
            var lStrResult = value;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        private bool CheckIsBlankSheet()
        {
            StopRow = _excelWorksheet.Dimension.End.Row;
            StopColumn = _excelWorksheet.Dimension.End.Column;
            for (var rowindex = 1; rowindex <= StopRow; rowindex++)
                for (var colIndex = 1; colIndex <= StopColumn; colIndex++)
                    if (_excelWorksheet.Cells[rowindex, colIndex].Value != null &&
                        _excelWorksheet.Cells[rowindex, colIndex].Text.Trim() != "")
                        return false;
            return true;
        }

        #endregion

        #endregion
    }
}