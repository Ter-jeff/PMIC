using OfficeOpenXml;
using PmicAutogen.InputPackages.Base;
using System.IO;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputScgh : ExcelInput
    {
        public InputScgh(FileInfo fileInfo) : base(fileInfo, InputFileType.ScghPatternList)
        {
        }

        protected override bool IsValidSheet(ExcelWorksheet sheet)
        {
            return true;
        }
    }
}