using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.OTP.Reader
{
    public class OtpDefaultRegList
    {
        public List<OtpDefaultRegRow> OtpDefaultRegRows { get; set; }

        public void Read(ExcelWorksheet sheet)
        {
            if (sheet == null) return;
            OtpDefaultRegRows = new List<OtpDefaultRegRow>();
            for (var row = 1; row <= sheet.Dimension.Rows; row++)
                if (sheet.Cells[row, 1].Value != null && sheet.Cells[row, 2].Value != null)
                {
                    var item = new OtpDefaultRegRow();
                    item.RegName = sheet.Cells[row, 1].Text;
                    item.RegAddress = sheet.Cells[row, 2].Text;
                    OtpDefaultRegRows.Add(item);
                }
        }

        public void ConvertRegisterMapToDefaultRegList(List<OtpRegister> otpRegisters)
        {
            OtpDefaultRegRows = new List<OtpDefaultRegRow>();
            var otpItems = otpRegisters.GroupBy(p => p.RegName).ToDictionary(p => p.Key, p => p.ToList());
            otpItems.Remove("");
            foreach (var otpItem in otpItems)
            {
                var otpDefaultRegRows = new OtpDefaultRegRow();
                otpDefaultRegRows.RegName = otpItem.Key;
                otpDefaultRegRows.RegAddress = ConvertRegisterValue(otpItem.Value.Select(p => p.OtpRegOfs).ToList(),
                    otpItem.Value.Select(p => p.DefaultValue).ToList());
                OtpDefaultRegRows.Add(otpDefaultRegRows);
            }
        }

        private string ConvertRegisterValue(List<string> offsets, List<string> values)
        {
            var result = "";
            var i = 0;
            foreach (var offset in offsets)
            {
                if (result.Length < int.Parse(offset)) result = result.PadLeft(int.Parse(offset), '0');
                result = ConvertDecimalToBinary(int.Parse(values[i])) + result;
                i++;
            }

            var strHex = Convert.ToInt32(result, 2).ToString("X2");
            return "0x" + strHex;
        }

        private string ConvertDecimalToBinary(int data)
        {
            if (data == 0) return "0";
            var result = "";
            while (data != 0)
            {
                int rem;
                data = Math.DivRem(data, 2, out rem);
                result = rem + result;
            }

            return result;
        }

        public class OtpDefaultRegRow
        {
            public string RegName { get; set; }
            public string RegAddress { get; set; }
        }
    }
}