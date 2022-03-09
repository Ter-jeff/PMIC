using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.AHBEnum.Input
{
    public class AhbRegisterMapRow
    {
        #region Constructor

        public AhbRegisterMapRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
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
            Comments = "";
        }

        #endregion

        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string Index { set; get; }
        public string Block { set; get; }
        public string RegAddress { set; get; }
        public string RegLink { set; get; }
        public string RegName { set; get; }
        public string FieldName { set; get; }
        public string FieldWidth { set; get; }
        public string FieldOffset { set; get; }
        public string FieldPosition { set; get; }
        public string FieldAccess { set; get; }
        public string FieldFormula { set; get; }
        public string IsDeterministic { set; get; }
        public string Type { set; get; }
        public string IsTestMode { set; get; }
        public string IsOtp { set; get; }
        public string OtpOwner { set; get; }
        public string OtpKey { set; get; }
        public string ReadBackValue { set; get; }
        public string HwResetValue { set; get; }
        public string Comments { set; get; }

        #endregion
    }

    public class AhbRegisterMapSheet
    {
        #region Constructor

        public AhbRegisterMapSheet(string name)
        {
            Name = name;
            Rows = new List<AhbRegisterMapRow>();
            HeaderIndexDic = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
        }

        #endregion

        public List<string> GenAhbEnum(bool fieldNameType, int maxBitWidth)
        {
            List<string> lines = new List<string>();
            int HexLength = maxBitWidth / 4;
            Dictionary<string, List<AhbRegisterMapRow>> regInfos = Rows.GroupBy(p => p.RegName + "#" + p.RegAddress)
                .ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<AhbRegisterMapRow>> regInfo in regInfos)
            {
                string regName = regInfo.Key.Split('#')[0];
                string regAddress = Regex
                    .Match(regInfo.Key.Split('#')[1], "0x(?<value>[a-f0-9]+)", RegexOptions.IgnoreCase).Groups["value"]
                    .ToString();
                lines.Add("Public Enum " + regName + "");
                lines.Add("    Addr = &H" + regAddress + "&");
                int reset = 0;
                foreach (AhbRegisterMapRow regItem in regInfo.Value)
                {
                    string fieldName = fieldNameType &&
                                       !regItem.FieldName.StartsWith(regItem.RegName, StringComparison.CurrentCulture)
                        ? regItem.RegName + "_" + regItem.FieldName
                        : regItem.FieldName;
                    int fieldWidth = int.Parse(regItem.FieldWidth);
                    int fieldOffset = int.Parse(regItem.FieldOffset);
                    string data = "".PadLeft(fieldWidth, '1') + "".PadLeft(fieldOffset, '0');
                    string dataHex = (~Convert.ToInt32(data, 2)).ToString("X8");
                    string address = dataHex.Substring(dataHex.Length - HexLength);
                    int hwResetValue = Convert.ToInt32(regItem.HwResetValue.Substring(2), 16);
                    reset += hwResetValue * (int)Math.Pow(2, fieldOffset);
                    lines.Add("    " + fieldName + " = &H" + address);
                }

                string resetHex = reset.ToString("X8");
                string resetAddress = resetHex.Substring(resetHex.Length - HexLength);
                lines.Add("    Default = &H" + resetAddress);
                lines.Add("End Enum");
            }

            return lines;
        }

        public int GetMaxBitWidth()
        {
            int max = 0;
            foreach (AhbRegisterMapRow row in Rows)
            {
                int fieldWidth;
                int fieldOffset;
                int.TryParse(row.FieldWidth, out fieldWidth);
                int.TryParse(row.FieldOffset, out fieldOffset);
                if (fieldWidth + fieldOffset > max)
                {
                    max = fieldWidth + fieldOffset;
                }
            }

            return max;
        }

        #region Properity

        public string Name { get; set; }
        public List<AhbRegisterMapRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndexDic { get; set; }

        #endregion
    }

    public class AhbRegisterMapReader
    {
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

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"index", true},
            {"block", true},
            {"reg address", true},
            {"reg link", true},
            {"reg name", true},
            {"field name", true},
            {"field width", true},
            {"field offset", true},
            {"field position", true},
            {"field access", true},
            {"field formula", true},
            {"isdeterministic", true},
            {"type", true},
            {"istestmode", true},
            {"isotp", true},
            {"otpowner", true},
            {"otpkey", true},
            {"readback value", true},
            {"hw reset value", true},
            {"comments", true}
        };

        private AhbRegisterMapSheet _ahbRegisterMapSheet;
        private int _blockIndex = -1;
        private int _commentsIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _fieldAccessIndex = -1;
        private int _fieldFormulaIndex = -1;
        private int _fieldNameIndex = -1;
        private int _fieldOffsetIndex = -1;
        private int _fieldPositionIndex = -1;
        private int _fieldWidthIndex = -1;
        private int _hwResetValueIndex = -1;
        private int _indexIndex = -1;
        private int _isDeterministicIndex = -1;
        private int _isOtpIndex = -1;
        private int _isTestModeIndex = -1;
        private int _otpKeyIndex = -1;
        private int _otpOwnerIndex = -1;
        private int _readBackValueIndex = -1;
        private int _regAddressIndex = -1;
        private int _regLinkIndex = -1;
        private int _regNameIndex = -1;
        private string _name;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public AhbRegisterMapSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _ahbRegisterMapSheet = new AhbRegisterMapSheet(_name);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _ahbRegisterMapSheet = ReadSheetData();

            return _ahbRegisterMapSheet;
        }

        private AhbRegisterMapSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                AhbRegisterMapRow row = new AhbRegisterMapRow(_name) { RowNum = i };
                if (_indexIndex != -1)
                {
                    row.Index = _excelWorksheet.GetMergeCellValue(i, _indexIndex).Trim();
                }

                if (_blockIndex != -1)
                {
                    row.Block = _excelWorksheet.GetMergeCellValue(i, _blockIndex).Trim();
                }

                if (_regAddressIndex != -1)
                {
                    row.RegAddress = _excelWorksheet.GetMergeCellValue(i, _regAddressIndex).Trim();
                }

                if (_regLinkIndex != -1)
                {
                    row.RegLink = _excelWorksheet.GetMergeCellValue(i, _regLinkIndex).Trim();
                }

                if (_regNameIndex != -1)
                {
                    row.RegName = _excelWorksheet.GetMergeCellValue(i, _regNameIndex).Trim();
                }

                if (_fieldNameIndex != -1)
                {
                    row.FieldName = _excelWorksheet.GetMergeCellValue(i, _fieldNameIndex).Trim();
                }

                if (_fieldWidthIndex != -1)
                {
                    row.FieldWidth = _excelWorksheet.GetMergeCellValue(i, _fieldWidthIndex).Trim();
                }

                if (_fieldOffsetIndex != -1)
                {
                    row.FieldOffset = _excelWorksheet.GetMergeCellValue(i, _fieldOffsetIndex).Trim();
                }

                if (_fieldPositionIndex != -1)
                {
                    row.FieldPosition = _excelWorksheet.GetMergeCellValue(i, _fieldPositionIndex).Trim();
                }

                if (_fieldAccessIndex != -1)
                {
                    row.FieldAccess = _excelWorksheet.GetMergeCellValue(i, _fieldAccessIndex).Trim();
                }

                if (_fieldFormulaIndex != -1)
                {
                    row.FieldFormula = _excelWorksheet.GetMergeCellValue(i, _fieldFormulaIndex).Trim();
                }

                if (_isDeterministicIndex != -1)
                {
                    row.IsDeterministic = _excelWorksheet.GetMergeCellValue(i, _isDeterministicIndex).Trim();
                }

                if (_typeIndex != -1)
                {
                    row.Type = _excelWorksheet.GetMergeCellValue(i, _typeIndex).Trim();
                }

                if (_isTestModeIndex != -1)
                {
                    row.IsTestMode = _excelWorksheet.GetMergeCellValue(i, _isTestModeIndex).Trim();
                }

                if (_isOtpIndex != -1)
                {
                    row.IsOtp = _excelWorksheet.GetMergeCellValue(i, _isOtpIndex).Trim();
                }

                if (_otpOwnerIndex != -1)
                {
                    row.OtpOwner = _excelWorksheet.GetMergeCellValue(i, _otpOwnerIndex).Trim();
                }

                if (_otpKeyIndex != -1)
                {
                    row.OtpKey = _excelWorksheet.GetMergeCellValue(i, _otpKeyIndex).Trim();
                }

                if (_readBackValueIndex != -1)
                {
                    row.ReadBackValue = _excelWorksheet.GetMergeCellValue(i, _readBackValueIndex).Trim();
                }

                if (_hwResetValueIndex != -1)
                {
                    row.HwResetValue = _excelWorksheet.GetMergeCellValue(i, _hwResetValueIndex).Trim();
                }

                if (_commentsIndex != -1)
                {
                    row.Comments = _excelWorksheet.GetMergeCellValue(i, _commentsIndex).Trim();
                }

                _ahbRegisterMapSheet.Rows.Add(row);
            }

            return _ahbRegisterMapSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderIndex, StringComparison.OrdinalIgnoreCase))
                {
                    _indexIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderIndex, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderBlock, StringComparison.OrdinalIgnoreCase))
                {
                    _blockIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderBlock, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderRegAddress, StringComparison.OrdinalIgnoreCase))
                {
                    _regAddressIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderRegAddress, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderRegLink, StringComparison.OrdinalIgnoreCase))
                {
                    _regLinkIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderRegLink, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderRegName, StringComparison.OrdinalIgnoreCase))
                {
                    _regNameIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderRegName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldNameIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldWidth, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldWidthIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldWidth, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldOffset, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldOffsetIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldOffset, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldPosition, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldPositionIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldPosition, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldAccess, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldAccessIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldAccess, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFieldFormula, StringComparison.OrdinalIgnoreCase))
                {
                    _fieldFormulaIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderFieldFormula, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderIsDeterministic, StringComparison.OrdinalIgnoreCase))
                {
                    _isDeterministicIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderIsDeterministic, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderType, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderIsTestMode, StringComparison.OrdinalIgnoreCase))
                {
                    _isTestModeIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderIsTestMode, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderIsOtp, StringComparison.OrdinalIgnoreCase))
                {
                    _isOtpIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderIsOtp, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOtpOwner, StringComparison.OrdinalIgnoreCase))
                {
                    _otpOwnerIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderOtpOwner, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderOtpKey, StringComparison.OrdinalIgnoreCase))
                {
                    _otpKeyIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderOtpKey, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderReadBackValue, StringComparison.OrdinalIgnoreCase))
                {
                    _readBackValueIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderReadBackValue, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderHwResetValue, StringComparison.OrdinalIgnoreCase))
                {
                    _hwResetValueIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderHwResetValue, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderComments, StringComparison.OrdinalIgnoreCase))
                {
                    _commentsIndex = i;
                    _ahbRegisterMapSheet.HeaderIndexDic.Add(HeaderComments, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _ahbRegisterMapSheet.HeaderIndexDic)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
                for (int j = 1; j <= colNum; j++)
                {
                    if (_excelWorksheet.GetMergeCellValue(i, j).Trim()
                        .Equals(HeaderIndex, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _indexIndex = -1;
            _blockIndex = -1;
            _regAddressIndex = -1;
            _regLinkIndex = -1;
            _regNameIndex = -1;
            _fieldNameIndex = -1;
            _fieldWidthIndex = -1;
            _fieldOffsetIndex = -1;
            _fieldPositionIndex = -1;
            _fieldAccessIndex = -1;
            _fieldFormulaIndex = -1;
            _isDeterministicIndex = -1;
            _typeIndex = -1;
            _isTestModeIndex = -1;
            _isOtpIndex = -1;
            _otpOwnerIndex = -1;
            _otpKeyIndex = -1;
            _readBackValueIndex = -1;
            _hwResetValueIndex = -1;
            _commentsIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> mappingDictionaries = new List<Dictionary<string, string>>();
            foreach (AhbRegisterMapRow row in _ahbRegisterMapSheet.Rows)
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string>
                {
                    {"index", row.Index},
                    {"block", row.Block},
                    {"reg address", row.RegAddress},
                    {"reg link", row.RegLink},
                    {"reg name", row.RegName},
                    {"field name", row.FieldName},
                    {"field width", row.FieldWidth},
                    {"field offset", row.FieldOffset},
                    {"field position", row.FieldPosition},
                    {"field access", row.FieldAccess},
                    {"field formula", row.FieldFormula},
                    {"isdeterministic", row.IsDeterministic},
                    {"type", row.Type},
                    {"istestmode", row.IsTestMode},
                    {"isotp", row.IsOtp},
                    {"otpowner", row.OtpOwner},
                    {"otpkey", row.OtpKey},
                    {"readback value", row.ReadBackValue},
                    {"hw reset value", row.HwResetValue},
                    {"comments", row.Comments}
                };
                mappingDictionaries.Add(dictionary);
            }

            return mappingDictionaries;
        }
    }
}