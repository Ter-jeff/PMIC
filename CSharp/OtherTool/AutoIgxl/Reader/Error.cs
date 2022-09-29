using System.Collections.Generic;

namespace AutoIgxl.Reader
{
    public class Error
    {
        private string _sheetName;

        public List<string> Comments = new List<string>();

        public string SheetName
        {
            get { return _sheetName.Length > 31 ? _sheetName.Substring(0, 31) : _sheetName; }
            set { _sheetName = value; }
        }

        public object ErrorType { set; get; }

        public string Link
        {
            get { return GetHyperlink(SheetName, RowNum, ColNum, "Link"); }
        }

        public ErrorLevel ErrorLevel { get; set; }
        public int RowNum { get; set; }
        public int ColNum { get; set; }
        public string Message { get; set; }

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