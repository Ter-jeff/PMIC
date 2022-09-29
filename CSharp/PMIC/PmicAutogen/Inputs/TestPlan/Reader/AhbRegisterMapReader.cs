using OfficeOpenXml;
using PmicAutogen.Config.ProjectConfig;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class AhbRegisterMapReader
    {
        //Index	Block	Reg Address	Reg Link	Reg Name	
        //Field Name	Field Width	Field Offset	Field Position	Field Access	Field Formula	
        //isDeterministic	Type	isTestMode	isOTP	OTPOwner	OTPKey	ReadBack Value	HW Initialize Value	Comments
        private const string HeaderIndex = "index";
        private const string HeaderBlock = "block";
        private const string HeaderRegAddress = "reg address";
        private const string HeaderRegLink = "reg link";
        private const string HeaderRegName = "reg name";
        private const string HeaderFieldName = "field name";
        private const string HeaderFieldWidth = "field width";
        private const string HeaderFieldOffset = "field offset";
        private const string HeaderFieldPosition = "field position";
        private const string HeaderFieldAccess = "field access";
        private const string HeaderFieldFormula = "field formula";
        private const string HeaderIsDeterministic = "isdeterministic";
        private const string HeaderType = "type";
        private const string HeaderIsTestMode = "istestmode";
        private const string HeaderIsOtp = "isotp";
        private const string HeaderOtpOwner = "otpowner";
        private const string HeaderOtpKey = "otpkey";
        private const string HeaderReadBackValue = "readback value";
        private const string HeaderHwResetValue = "hw reset value";
        private const string HeaderComments = "comments";
        private readonly Dictionary<string, int> _headersDictionary = new Dictionary<string, int>();

        public AhbRegisterMapSheet ReadSheet(ExcelWorksheet sheet)
        {
            if (sheet == null) return null;
            var ahbRegSheet = new AhbRegisterMapSheet();
            ahbRegSheet.Name = sheet.Name;
            var rowIndex = 1;
            for (var row = 1; row <= sheet.Dimension.Rows; row++)
            {
                for (var col = 1; col <= sheet.Dimension.Columns; col++)
                    if (ahbRegSheet.Headers.Contains(sheet.Cells[row, col].Text.ToLower()))
                        _headersDictionary.Add(sheet.Cells[row, col].Text.ToLower(), col);
                rowIndex++;
                if (_headersDictionary.Count > 0) break;
            }

            if (_headersDictionary.Count == 0) return null;

            for (var row = rowIndex; row <= sheet.Dimension.Rows; row++)
            {
                var item = new AhbRegRow();
                foreach (var header in _headersDictionary)
                {
                    if (header.Key == HeaderIndex) item.Index = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderBlock) item.Block = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderRegAddress) item.RegAddress = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderRegLink) item.RegLink = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderRegName) item.RegName = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldName) item.FieldName = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldWidth) item.FieldWidth = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldOffset) item.FieldOffset = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldPosition) item.FieldPosition = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldAccess) item.FieldAccess = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderFieldFormula) item.FieldFormula = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderIsDeterministic) item.IsDeterministic = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderType) item.Type = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderIsTestMode) item.IsTestMode = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderIsOtp) item.IsOtp = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderOtpOwner) item.OtpOwner = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderOtpKey) item.OtpKey = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderReadBackValue) item.ReadBackValue = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderHwResetValue) item.HwResetValue = sheet.Cells[row, header.Value].Text;
                    if (header.Key == HeaderComments) item.Comment = sheet.Cells[row, header.Value].Text;
                }

                ahbRegSheet.AhbRegRows.Add(item);
            }

            return ahbRegSheet;
        }
    }

    public class AhbRegisterMapSheet
    {
        public List<AhbRegRow> AhbRegRows = new List<AhbRegRow>();

        public List<string> Headers = new List<string>
        {
            "index",
            "block",
            "reg address",
            "reg link",
            "reg name",
            "field name",
            "field width",
            "field offset",
            "field position",
            "field access",
            "field formula",
            "isdeterministic",
            "type",
            "istestmode",
            "isotp",
            "otpowner",
            "otpkey",
            "readback value",
            "hw reset value",
            "comments"
        };

        public string Name { get; set; }


        public List<string> WriteAhbEnum(string dir)
        {
            var ahbEnumName = ProjectConfigSingleton.Instance().GetProjectConfigValue("OTP", "AHBEnumName");
            var fileList = new List<string>();
            var regInfos = AhbRegRows.GroupBy(p => p.RegName + "#" + p.RegAddress)
                .ToDictionary(p => p.Key, p => p.ToList());
            var ahbFileCount = 1;
            var index = 0;
            var filename = string.Format("AHB_REG_MAP{0}", ahbFileCount);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var writer = new StreamWriter(Path.Combine(dir, string.Format("{0}.bas", filename)));
            writer.WriteLine("Attribute VB_Name = \"{0}\"", filename);
            foreach (var regInfo in regInfos)
            {
                var regName = regInfo.Key.Split('#')[0];
                if (string.IsNullOrEmpty(regName)) continue;
                var regAddress = Regex
                    .Match(regInfo.Key.Split('#')[1], "0x(?<value>[a-f0-9]+)", RegexOptions.IgnoreCase).Groups["value"]
                    .ToString();
                string name;
                if (ahbEnumName.Equals("reg name", StringComparison.CurrentCultureIgnoreCase))
                {
                    name = regInfo.Value[0].RegName;
                }
                else
                {
                    if (!regInfo.Value[0].RegName
                            .EndsWith(regInfo.Value[0].FieldName, StringComparison.CurrentCultureIgnoreCase))
                        name = regInfo.Value[0].RegName + "_" + regInfo.Value[0].FieldName;
                    else
                        name = regInfo.Value[0].RegName;
                }

                writer.WriteLine("'Public Enum {0}", name);
                writer.WriteLine("'Addr = &H{0}&", regAddress);
                foreach (var regItem in regInfo.Value)
                {
                    var fieldName = regItem.FieldName;
                    var fieldWidth = int.Parse(regItem.FieldWidth);
                    var fieldOffset = int.Parse(regItem.FieldOffset);
                    var data = string.Format("{0}{1}", "".PadLeft(fieldWidth, '1'), "".PadLeft(fieldOffset, '0'));
                    var dataHex = (~Convert.ToInt32(data, 2)).ToString("X4");

                    var address = dataHex.Substring(dataHex.Length - 2, 2);
                    if (address.StartsWith("0"))
                        address = address.Substring(1, 1);
                    writer.WriteLine("'{0} = &H{1}", fieldName, address);
                }

                writer.WriteLine("'End Enum");
                index++;
                if (index > 1000)
                {
                    fileList.Add(Path.Combine(dir, string.Format("{0}.bas", filename)));
                    index = 0;
                    writer.Close();
                    ahbFileCount++;
                    filename = string.Format("AHB_REG_MAP{0}", ahbFileCount);
                    writer = new StreamWriter(Path.Combine(dir, string.Format("{0}.bas", filename)));
                    writer.WriteLine("Attribute VB_Name = \"{0}\"", filename);
                }
            }

            writer.Close();
            fileList.Add(Path.Combine(dir, string.Format("{0}.bas", filename)));
            return fileList;
        }
    }

    public class AhbRegRow
    {
        private string _block;
        private string _comment;
        private string _fieldAccess;
        private string _fieldFormula;
        private string _fieldName;
        private string _fieldOffset;
        private string _fieldPosition;
        private string _fieldWidth;
        private string _hwResetValue;
        private string _index;
        private string _isDeterministic;
        private string _isOtp;
        private string _isTestMode;
        private string _otpKey;
        private string _readBackValue;
        private string _regAddress;
        private string _regName;
        private string _type;

        public AhbRegRow()
        {
            Index = "";
            Block = "";
            RegAddress = "";
            RegLink = "";
            RegName = "";
            FieldName = "";
            FieldWidth = "";
            FieldOffset = "";
            FieldPosition = "";
            FieldAccess = "";
            FieldFormula = "";
            IsDeterministic = "";
            Type = "";
            IsTestMode = "";
            IsOtp = "";
            OtpOwner = "";
            OtpKey = "";
            ReadBackValue = "";
            HwResetValue = "";
            Comment = "";
        }

        public string Index
        {
            get { return _index.ToUpper(); }
            set { _index = value; }
        }

        public string Block
        {
            get { return _block.ToUpper(); }
            set { _block = value; }
        }

        public string RegAddress
        {
            get { return _regAddress.ToUpper(); }
            set { _regAddress = value; }
        }

        public string RegLink { get; set; }

        public string RegName
        {
            get { return _regName.ToUpper(); }
            set { _regName = value; }
        }

        public string FieldName
        {
            get { return _fieldName.ToUpper(); }
            set { _fieldName = value; }
        }

        public string FieldWidth
        {
            get { return _fieldWidth.ToUpper(); }
            set { _fieldWidth = value; }
        }

        public string FieldOffset
        {
            get { return _fieldOffset.ToUpper(); }
            set { _fieldOffset = value; }
        }

        public string FieldPosition
        {
            get { return _fieldPosition.ToUpper(); }
            set { _fieldPosition = value; }
        }

        public string FieldAccess
        {
            get { return _fieldAccess.ToUpper(); }
            set { _fieldAccess = value; }
        }

        public string FieldFormula
        {
            get { return _fieldFormula.ToUpper(); }
            set { _fieldFormula = value; }
        }

        public string IsDeterministic
        {
            get { return _isDeterministic.ToUpper(); }
            set { _isDeterministic = value; }
        }

        public string Type
        {
            get { return _type.ToUpper(); }
            set { _type = value; }
        }

        public string IsTestMode
        {
            get { return _isTestMode.ToUpper(); }
            set { _isTestMode = value; }
        }

        public string IsOtp
        {
            get { return _isOtp.ToUpper(); }
            set { _isOtp = value; }
        }

        public string OtpOwner { get; set; }

        public string OtpKey
        {
            get { return _otpKey.ToUpper(); }
            set { _otpKey = value; }
        }

        public string ReadBackValue
        {
            get { return _readBackValue.ToUpper(); }
            set { _readBackValue = value; }
        }

        public string HwResetValue
        {
            get { return _hwResetValue.ToUpper(); }
            set { _hwResetValue = value; }
        }

        public string Comment
        {
            get { return _comment.ToUpper(); }
            set { _comment = value; }
        }

        public List<string> GetRegData()
        {
            return new List<string>
            {
                Index,
                Block,
                RegAddress,
                RegLink,
                RegName,
                FieldName,
                FieldWidth,
                FieldOffset,
                FieldPosition,
                FieldAccess,
                FieldFormula,
                IsDeterministic,
                Type,
                IsTestMode,
                IsOtp,
                OtpOwner,
                OtpKey,
                ReadBackValue,
                HwResetValue,
                Comment
            };
        }
    }
}