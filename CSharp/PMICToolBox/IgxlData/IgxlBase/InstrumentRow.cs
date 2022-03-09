using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace IgxlData.IgxlBase
{
    public class InstrumentRow
    {
        private string _min;
        private string _max;
        private string _currentType;
        private string _isMerge;
        private string _instrument;
        private string _iFold;

        public string Min
        {
            get { return _min; }
            set { _min = value; }
        }
        public string Max
        {
            get { return _max; }
            set { _max = value; }
        }
        public string CurrentType
        {
            get { return _currentType; }
            set { _currentType = value; }
        }
        public string IsMerge
        {
            get { return _isMerge; }
            set { _isMerge = value; }
        }
        public string Instrument
        {
            get { return _instrument; }
            set { _instrument = value; }
        }
        public string IFold
        {
            get { return _iFold; }
            set { _iFold = value; }
        }

        public InstrumentRow(string min, string max, string currentType, string isMerge, string instrument, string iFold)
        {
            _min = min;
            _max = max;
            _currentType = currentType;
            _isMerge = isMerge;
            _instrument = instrument;
            _iFold = iFold;
        }

        public static List<InstrumentRow> Reader(string filename)
        {
            List<InstrumentRow> listInstRow = new List<InstrumentRow>();
            if (!File.Exists(filename)) return listInstRow;
            ExcelWorksheet worksheet = null;
            Dictionary<string, int> dicHeader = new Dictionary<string, int>();
            string min = "";
            string max = "";
            string currenttype = "";
            string ismerge = "";
            string instrument = "";
            string ifold = "";

            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    worksheet = ep.Workbook.Worksheets["IFold"];
                    if (worksheet == null) return listInstRow;

                    // get header index
                    for (int col = 1; col <= worksheet.Dimension.End.Column; ++col)
                    {
                        string header = worksheet.Cells[1, col].Value != null ? worksheet.Cells[1, col].Value.ToString().Trim() : "";
                        if (Regex.IsMatch(header, @"Min", RegexOptions.IgnoreCase)) dicHeader.Add("Min", 1);
                        else if (Regex.IsMatch(header, @"Max", RegexOptions.IgnoreCase)) dicHeader.Add("Max", 2);
                        else if (Regex.IsMatch(header, @"CurrentType", RegexOptions.IgnoreCase)) dicHeader.Add("CurrentType", 3);
                        else if (Regex.IsMatch(header, @"IsMerge", RegexOptions.IgnoreCase)) dicHeader.Add("IsMerge", 4);
                        else if (Regex.IsMatch(header, @"Instrument", RegexOptions.IgnoreCase)) dicHeader.Add("Instrument", 5);
                        else if (Regex.IsMatch(header, @"Ifold", RegexOptions.IgnoreCase)) dicHeader.Add("Ifold", 6);
                    }
                    if (dicHeader.Count < 6) return listInstRow;

                    for (int row = 2; row <= worksheet.Dimension.End.Row; ++row)
                    {
                        min = worksheet.Cells[row, dicHeader["Min"]].Value.ToString();
                        max = worksheet.Cells[row, dicHeader["Max"]].Value.ToString();
                        currenttype = worksheet.Cells[row, dicHeader["CurrentType"]].Value != null ? worksheet.Cells[row, dicHeader["CurrentType"]].Value.ToString() : "";
                        ismerge = worksheet.Cells[row, dicHeader["IsMerge"]].Value.ToString();
                        instrument = worksheet.Cells[row, dicHeader["Instrument"]].Value.ToString();
                        ifold = worksheet.Cells[row, dicHeader["Ifold"]].Value.ToString();
                        listInstRow.Add(new InstrumentRow(min, max, currenttype, ismerge, instrument, ifold));
                    }
                }
            }
            return listInstRow;
        }
    }
}
