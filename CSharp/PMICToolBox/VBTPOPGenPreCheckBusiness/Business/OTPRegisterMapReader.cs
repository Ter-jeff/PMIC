using VBTPOPGenPreCheckBusiness.DataStore;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace VBTPOPGenPreCheckBusiness.Business
{
    public class OTPRegisterMapReader
    {
        private string _file;
        private string otpRegName = "OTP_REGISTER_NAME";
        private string otpOwner = "otp_owner";
        private string regName = "reg_name";
        private string name = "name";
        private string bw = "bw";
        private int otpRegIndex = -1;
        private int otpOwnerIndex = -1;
        private int regNameIndex = -1;
        private int nameIndex = -1;
        private int bwIndex = -1;
        private int headerRowIndex = -1;
        private Dictionary<string, OTPRegisterMapData> _dicOTPRegMap; // key: OTP_REGISTER_NAME, value: otp_owner

        public Dictionary<string, OTPRegisterMapData> dicOTPRegMap
        {
            get { return _dicOTPRegMap; }
        }

        public OTPRegisterMapReader(string file)
        {
            if (!File.Exists(file))
                return;
            _file = file;
            _dicOTPRegMap = new Dictionary<string, OTPRegisterMapData>();
            Read();
        }

        public void Read()
        {
            DataTable regMapDT = Library.Common.Utility.ConvertToDataTable(_file, new char[] { '\t' }, otpRegName);
            FetchInfo(regMapDT);
            for (int r = headerRowIndex + 1; r < regMapDT.Rows.Count; ++r)
            {
                string otp_reg_name = regMapDT.Rows[r][otpRegIndex].ToString();
                if (otp_reg_name.ToUpper().Equals("END"))
                    break;

                string otp_owner = regMapDT.Rows[r][otpOwnerIndex].ToString();
                string regName = regMapDT.Rows[r][regNameIndex].ToString();
                string name = regMapDT.Rows[r][nameIndex].ToString();
                string bw = regMapDT.Rows[r][bwIndex].ToString();

                if (!_dicOTPRegMap.ContainsKey(otp_reg_name))
                    _dicOTPRegMap.Add(otp_reg_name, new OTPRegisterMapData(otp_owner, regName, name, bw));
            }
        }

        private void FetchInfo(DataTable regMapDT)
        {
            int rowIndex = -1;
            foreach (DataRow row in regMapDT.Rows)
            {
                ++rowIndex;
                int colIndex = -1;
                foreach (var item in row.ItemArray)
                {
                    ++colIndex;
                    if (item == null || item.ToString().Equals(string.Empty))
                        break;

                    if (item.ToString().Equals(otpRegName, StringComparison.OrdinalIgnoreCase))
                    {
                        headerRowIndex = rowIndex;
                        otpRegIndex = colIndex;
                    }
                    else if (item.ToString().Equals(otpOwner, StringComparison.OrdinalIgnoreCase))
                        otpOwnerIndex = colIndex;
                    else if (item.ToString().Equals(regName, StringComparison.OrdinalIgnoreCase))
                        regNameIndex = colIndex;
                    else if (item.ToString().Equals(name, StringComparison.OrdinalIgnoreCase))
                        nameIndex = colIndex;
                    else if (item.ToString().Equals(bw, StringComparison.OrdinalIgnoreCase))
                        bwIndex = colIndex;

                    if (otpRegIndex != -1 && otpOwnerIndex != -1 && regNameIndex != -1 && nameIndex != -1 && bwIndex != -1)
                        break;
                }
                if (otpRegIndex != -1 && otpOwnerIndex != -1 && regNameIndex != -1 && nameIndex != -1 && bwIndex != -1)
                    break;
            }
        }
    }
}
