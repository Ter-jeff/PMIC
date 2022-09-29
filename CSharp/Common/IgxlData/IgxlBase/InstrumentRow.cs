using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    public class InstrumentRow
    {
        public InstrumentRow(string min, string max, string currentType, string isMerge, string instrument,
            string iFold)
        {
            Min = min;
            Max = max;
            CurrentType = currentType;
            IsMerge = isMerge;
            Instrument = instrument;
            Fold = iFold;
        }

        public string Min { get; set; }

        public string Max { get; set; }

        public string CurrentType { get; set; }

        public string IsMerge { get; set; }

        public string Instrument { get; set; }

        public string Fold { get; set; }

        public static List<InstrumentRow> Reader(string filename)
        {
            var listInstRow = new List<InstrumentRow>();
            if (!File.Exists(filename)) return listInstRow;
            var dicHeader = new Dictionary<string, int>();

            using (var fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var ep = new ExcelPackage(fs))
                {
                    var worksheet = ep.Workbook.Worksheets["IFold"];
                    if (worksheet == null) return listInstRow;

                    // get header index
                    for (var col = 1; col <= worksheet.Dimension.End.Column; ++col)
                    {
                        var header = worksheet.Cells[1, col].Value != null
                            ? worksheet.Cells[1, col].Value.ToString().Trim()
                            : "";
                        if (Regex.IsMatch(header, @"Min", RegexOptions.IgnoreCase)) dicHeader.Add("Min", 1);
                        else if (Regex.IsMatch(header, @"Max", RegexOptions.IgnoreCase)) dicHeader.Add("Max", 2);
                        else if (Regex.IsMatch(header, @"CurrentType", RegexOptions.IgnoreCase))
                            dicHeader.Add("CurrentType", 3);
                        else if (Regex.IsMatch(header, @"IsMerge", RegexOptions.IgnoreCase))
                            dicHeader.Add("IsMerge", 4);
                        else if (Regex.IsMatch(header, @"Instrument", RegexOptions.IgnoreCase))
                            dicHeader.Add("Instrument", 5);
                        else if (Regex.IsMatch(header, @"Ifold", RegexOptions.IgnoreCase)) dicHeader.Add("Ifold", 6);
                    }

                    if (dicHeader.Count < 6) return listInstRow;

                    for (var row = 2; row <= worksheet.Dimension.End.Row; ++row)
                    {
                        var min = worksheet.Cells[row, dicHeader["Min"]].Value.ToString();
                        var max = worksheet.Cells[row, dicHeader["Max"]].Value.ToString();
                        var currentType = worksheet.Cells[row, dicHeader["CurrentType"]].Value != null
                            ? worksheet.Cells[row, dicHeader["CurrentType"]].Value.ToString()
                            : "";
                        var isMerge = worksheet.Cells[row, dicHeader["IsMerge"]].Value.ToString();
                        var instrument = worksheet.Cells[row, dicHeader["Instrument"]].Value.ToString();
                        var ifold = worksheet.Cells[row, dicHeader["Ifold"]].Value.ToString();
                        listInstRow.Add(new InstrumentRow(min, max, currentType, isMerge, instrument, ifold));
                    }
                }
            }

            return listInstRow;
        }
    }
}