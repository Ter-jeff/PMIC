using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.IgxlSheets
{
    public class OtherSheet : IgxlSheet
    {
        public List<string> Lines = new List<string>();

        public OtherSheet(Worksheet sheet) : base(sheet)
        {
        }

        public OtherSheet(ExcelWorksheet sheet) : base(sheet)
        {
        }

        public OtherSheet(string sheetName) : base(sheetName)
        {
        }

        public override void Write(string fileName, string version = "")
        {
            using (var sw = new StreamWriter(fileName))
            {
                foreach (var line in Lines)
                    sw.WriteLine(line);
            }
        }
    }
}