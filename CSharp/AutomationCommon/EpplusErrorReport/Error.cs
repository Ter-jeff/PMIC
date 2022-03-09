using System.Collections.Generic;

namespace AutomationCommon.EpplusErrorReport
{
    public class Error
    {
        public string ErrorType { set; get; }
        public ErrorLevel ErrorLevel { get; set; }
        private string _sheetName;
        public string SheetName
        {
            get { return _sheetName.Length > 31 ? _sheetName.Substring(0, 31) : _sheetName; }
            set { _sheetName = value; }
        }
        public string Link
        {
            get
            {
                return GetHyperlink(SheetName, RowNum, ColNum, "Link");
            }
            set { throw new System.NotImplementedException(); }
        }

        public int RowNum { get; set; }
        public int ColNum { get; set; }
        public string Message { get; set; }
        public List<string> Comments;
      

        private string GetHyperlink(string sheetName, int row, int column, string friendlyName)
        {
            if (column == 0)
                return "=HYPERLINK(\"#\'" + sheetName + "\'!" + row + ":" + row + "\",\"" + friendlyName + "\")";
            return "=HYPERLINK(\"#\'" + sheetName + "\'!" + GetAddress(row, column) + "\",\"" + friendlyName + "\")";
        }

        private string GetAddress(int row, int column, bool absolute = false)
        {
            if (row == 0 || column == 0)
                return "#REF!";
            if (absolute)
                return ("$" + GetColumnLetter(column) + "$" + row);
            return (GetColumnLetter(column) + row);
        }

        private string GetColumnLetter(int iColumnNumber)
        {
            if (iColumnNumber < 1)
                return "#REF!";

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return sCol;
        }
    }
}