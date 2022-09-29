using CommonLib.Enum;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace CommonLib.ErrorReport
{
    [DebuggerDisplay("{Message}")]
    public class Error
    {
        private string _sheetName = "";
        public List<string> Comments;
        public EnumErrorLevel ErrorLevel;
        public EnumErrorType EnumErrorType;
        public string _errorType;
        public string ErrorType
        {
            get
            {
                if (!string.IsNullOrEmpty(_errorType))
                    return _errorType;
                return EnumErrorType.ToString();
            }
            set { _errorType = value; } //Need Set SetOutline<T>
        }

        public string Level
        {
            get { return ErrorLevel.ToString(); }
        }

        public string SheetName
        {
            get { return _sheetName.Length > 31 ? _sheetName.Substring(0, 31) : _sheetName; }
            set { _sheetName = value; }
        }

        public string Link
        {
            get { return GetHyperlink(); }
            set { throw new NotImplementedException(); }
        }

        public int RowNum { get; set; }
        public int ColNum { get; set; }
        public string Message { get; set; }

        public string GetHyperlink()
        {
            if (string.IsNullOrEmpty(SheetName))
                return "";

            var friendlyName = "Link";
            if (ColNum == 0)
                return "=HYPERLINK(\"#\'" + SheetName + "\'!" + RowNum + ":" + RowNum + "\",\"" + friendlyName + "\")";
            return "=HYPERLINK(\"#\'" + SheetName + "\'!" + GetAddress(RowNum, ColNum) + "\",\"" + friendlyName + "\")";
        }

        private string GetAddress(int row, int column, bool absolute = false)
        {
            if (row == 0 || column == 0)
                return "#REF!";
            if (absolute)
                return "$" + GetColumnLetter(column) + "$" + row;
            return GetColumnLetter(column) + row;
        }

        private string GetColumnLetter(int iColumnNumber)
        {
            if (iColumnNumber < 1)
                return "#REF!";

            var sCol = "";
            do
            {
                sCol = (char)('A' + (iColumnNumber - 1) % 26) + sCol;
                iColumnNumber = (iColumnNumber - (iColumnNumber - 1) % 26) / 26;
            } while (iColumnNumber > 0);

            return sCol;
        }
    }
}