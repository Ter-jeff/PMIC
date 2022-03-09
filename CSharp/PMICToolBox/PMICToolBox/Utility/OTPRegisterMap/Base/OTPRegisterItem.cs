using System.Collections.Generic;

namespace PmicAutomation.Utility.OTPRegisterMap.Base
{

    public class OtpRegisterItem
    {
        //OTP_REGISTER_NAME	name	inst_name	reg_name	otp_owner	
        //DEFAULT VALUE	bw	idx	offset	otp_b0	otp_a0	otpreg_add	otpreg_ofs	Default or Real	REAL VALUE	Comment
        public string OtpRegisterName { get; set; }
        public string Name { get; set; }
        public string InstName { get; set; }
        public string InstBase { get; set; }
        public string RegName { get; set; }
        public string RegOfs { get; set; }
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
        public string RealValue { get; set; }
        public string Comment { get; set; }
        public List<string> OtpExtra { get; set; }

        public OtpRegisterItem()
        {
            OtpExtra = new List<string>();
        }
    }
}