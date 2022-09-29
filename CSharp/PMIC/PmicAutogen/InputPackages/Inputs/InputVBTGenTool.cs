using OfficeOpenXml;
using PmicAutogen.InputPackages.Base;
using System;
using System.IO;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputVbtGenTool : ExcelInput
    {
        public InputVbtGenTool(FileInfo fileInfo) : base(fileInfo, InputFileType.VbtGenTool)
        {
        }

        protected override bool IsValidSheet(ExcelWorksheet sheet)
        {
            if (sheet.Name.EndsWith("_TestParameter", StringComparison.CurrentCultureIgnoreCase))
                return true;

            //if (sheet.Cells[1, 2] == null || sheet.Cells[1, 2].Value == null) return false;
            //var name = sheet.Cells[1, 2].Value.ToString();
            //if (Regex.IsMatch(name, @"TEST PLAN FOR <", RegexOptions.IgnoreCase) && Regex.IsMatch(sheet.Name, @"_(?<Sequence>\d+)$", RegexOptions.IgnoreCase))
            //    return true;

            return false;
        }
    }
}