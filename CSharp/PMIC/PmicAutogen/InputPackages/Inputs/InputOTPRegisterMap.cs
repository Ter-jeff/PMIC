using PmicAutogen.InputPackages.Base;
using System.IO;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputOtpRegisterMap : Input
    {
        public InputOtpRegisterMap(FileInfo fileInfo) : base(fileInfo, InputFileType.OtpRegisterMap)
        {
        }
    }
}