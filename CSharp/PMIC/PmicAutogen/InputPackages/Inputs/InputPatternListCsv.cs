using PmicAutogen.InputPackages.Base;
using System.IO;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputPatternListCsv : Input
    {
        public InputPatternListCsv(FileInfo fileInfo) : base(fileInfo, InputFileType.PatternListCsv)
        {
        }
    }
}