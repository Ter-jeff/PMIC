using System.IO;
using PmicAutogen.InputPackages.Base;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputPatternListCsv : Input
    {
        public InputPatternListCsv(FileInfo fileInfo) : base(fileInfo, InputFileType.PatternListCsv)
        {
        }
    }
}