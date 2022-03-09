using System.IO;
using PmicAutogen.InputPackages.Base;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputOtpRegisterMap : Input
    {
        public InputOtpRegisterMap(FileInfo fileInfo) : base(fileInfo, InputFileType.OtpRegisterMap)
        {
        }
    }
}