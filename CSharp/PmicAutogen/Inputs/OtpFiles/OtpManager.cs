using System.Linq;
using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.Local;

namespace PmicAutogen.Inputs.OtpFiles
{
    public class OtpManager
    {
        public OtpFileReader OtpFileReader;

        public void CheckAll()
        {
            #region Pre check

            OtpFileReader = new OtpFileReader(LocalSpecs.YamlFileName);
            if (LocalSpecs.OtpFileName != null && LocalSpecs.OtpFileName.Any())
                foreach (var otpFileName in LocalSpecs.OtpFileName)
                {
                    var otpReader = new OtpFileReader(otpFileName);
                    OtpFileReader.MergeToYaml(otpReader.OtProws, otpReader.Version);
                }

            #endregion

            #region Post check

            #endregion
        }
    }
}