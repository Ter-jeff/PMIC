using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBTPOPGenPreCheckBusiness.DataStore
{
    public class OTPRegisterMapData
    {
        private string _otp_owner;
        private string _regName;
        private string _name;
        private string _bw;

        public OTPRegisterMapData(string otp_owner, string regName, string name, string bw)
        {
            _otp_owner = otp_owner;
            _regName = regName;
            _name = name;
            _bw = bw;
        }

        public string otp_owner
        {
            get { return _otp_owner; }
        }

        public string regName
        {
            get { return _regName; }
        }

        public string name
        {
            get { return _name; }
        }

        public string bw
        {
            get { return _bw; }
        }
    }
}
