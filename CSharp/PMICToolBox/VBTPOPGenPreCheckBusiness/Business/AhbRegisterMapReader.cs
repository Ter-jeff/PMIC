using System.Collections.Generic;
using System.IO;
using System.Data;

namespace VBTPOPGenPreCheckBusiness.Business
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

        public AhbRegisterMapSheet Read(string fileName)
        {
            if (!File.Exists(fileName)) return null;
            DataTable regMapDT = Library.Common.Utility.ConvertToDataTable(fileName, new char[] { '\t' }, HeaderIndex);
            var ahbRegSheet = new AhbRegisterMapSheet();
            ahbRegSheet.Name = fileName;
            var rowIndex = 0;
            for (var row = 0; row < regMapDT.Rows.Count; row++)
            {
                for (var col = 0; col < regMapDT.Columns.Count; col++)
                    if (ahbRegSheet.Headers.Contains(regMapDT.Rows[row][col].ToString().ToLower()))
                        _headersDictionary.Add(regMapDT.Rows[row][col].ToString().ToLower(), col);
                rowIndex++;
                if (_headersDictionary.Count > 0) break;
            }

            if (_headersDictionary.Count == 0) return null;

            for (var row = rowIndex; row < regMapDT.Rows.Count; row++)
            {
                var item = new AhbRegRow();
                foreach (var header in _headersDictionary)
                {
                    if (header.Key == HeaderIndex) item.Index = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderBlock) item.Block = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderRegAddress) item.RegAddress = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderRegLink) item.RegLink = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderRegName) item.RegName = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldName) item.FieldName = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldWidth) item.FieldWidth = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldOffset) item.FieldOffset = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldPosition) item.FieldPosition = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldAccess) item.FieldAccess = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderFieldFormula) item.FieldFormula = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderIsDeterministic) item.IsDeterministic = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderType) item.Type = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderIsTestMode) item.IsTestMode = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderIsOtp) item.IsOtp = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderOtpOwner) item.OtpOwner = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderOtpKey) item.OtpKey = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderReadBackValue) item.ReadBackValue = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderHwResetValue) item.HwResetValue = regMapDT.Rows[row][header.Value].ToString();
                    if (header.Key == HeaderComments) item.Comment = regMapDT.Rows[row][header.Value].ToString();
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