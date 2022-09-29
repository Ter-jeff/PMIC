using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.OtpFiles
{
    public class OtpManager
    {
        public OtpFileReader OtpFileReader;

        public void CheckAll()
        {
            #region Pre check

            OtpFileReader = new OtpFileReader(LocalSpecs.YamlFileName);
            if (LocalSpecs.OtpFileNames != null && LocalSpecs.OtpFileNames.Any())
            {
                var otpFileNames = new List<string> { LocalSpecs.YamlFileName };
                otpFileNames.AddRange(LocalSpecs.OtpFileNames);
                foreach (var otpFileName in otpFileNames)
                {
                    var otpReader = new OtpFileReader(otpFileName);
                    OtpFileReader.MergeToYaml(otpReader.OtProws, otpReader.Version);
                }
            }

            #endregion

            #region Post check

            #endregion
        }
    }
}