using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.OTP.Reader
{
    public class OtpRegister
    {
        public OtpRegister()
        {
            OtpExtra = new List<string>();
        }

        public string OtpRegisterName { get; set; }
        public string Name { get; set; }
        public string InstName { get; set; }
        public string InstBase { get; set; }
        public string RegOfs { get; set; }
        public string RegName { get; set; }
        public string OtpOwner { get; set; }
        public string DefaultValue { get; set; }
        public string Bw { get; set; }
        public string Idx { get; set; }
        public string Offset { get; set; }
        public string OtpB0 { get; set; }
        public string OtpA0 { get; set; }
        public string OtpRegAdd { get; set; }
        public string OtpRegOfs { get; set; }
        public string DefaultOrReal { get; set; }
        public string Comment { get; set; }
        public List<string> OtpExtra { get; set; }
    }
}