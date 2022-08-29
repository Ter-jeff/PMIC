using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.IgxlSheets
{
    public class OtherSheet : IgxlSheet
    {
        public OtherSheet(Worksheet sheet) : base(sheet)
        {
        }

        public OtherSheet(ExcelWorksheet sheet) : base(sheet)
        {
        }

        public OtherSheet(string sheetName) : base(sheetName)
        {
        }

        public List<string> Lines = new List<string>();

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
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