using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.Inputs.OtpFiles;

namespace PmicAutogen.Local
{
    public static class StaticOtp
    {
        public static OtpFileReader OtpFileReader;

        public static void AddSheets(OtpManager otpManager)
        {
            OtpFileReader = otpManager.OtpFileReader;
        }
    }
}