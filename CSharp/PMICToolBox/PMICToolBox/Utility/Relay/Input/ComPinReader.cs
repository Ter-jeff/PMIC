using PmicAutomation.Utility.Relay.Base;
using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.Relay.Input
{
    public class ComPinRow
    {
        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string Refdes { set; get; }
        public int PinNumber { set; get; }
        public string CompDeviceType { set; get; }
        public string PinType { set; get; }
        public string PinName { set; get; }
        public string NetName { set; get; }

        #endregion

        #region Constructor

        public ComPinRow()
        {
            Refdes = "";
            CompDeviceType = "";
            PinType = "";
            PinName = "";
            NetName = "";
        }

        public ComPinRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Refdes = "";
            CompDeviceType = "";
            PinType = "";
            PinName = "";
            NetName = "";
        }

        #endregion
    }

    public class ComPinSheet
    {
        #region Constructor

        public ComPinSheet(string name)
        {
            Name = name;
            Rows = new List<ComPinRow>();
        }

        #endregion

        public List<ComPinRow> FilterRows(LinkedNodeRuleSheet linkedNodeRuleSheet,
            Dictionary<string, List<string>> filterPinDic)
        {
            //Excluding by table
            List<ComPinRow> rows = Rows;
            //if (linkedNodeRuleSheet != null)
            //{
            //    List<string> nodelist = linkedNodeRuleSheet.Rows.Select(x => x.Node.ToUpper()).ToList();
            //    List<string> linkedNodelist = linkedNodeRuleSheet.Rows.Select(x => x.LinkedNode.ToUpper()).ToList();
            //    nodelist.AddRange(linkedNodelist);
            //    rows = Rows.Where(x => nodelist.Contains(x.PinName.ToUpper())).ToList();
            //}

            //Excluding not startwiths S0
            //rows = rows.Where(x =>  x.Refdes.StartsWith("S0",StringComparison.CurrentCultureIgnoreCase)).ToList();

            //Excluding duplicate
            rows = rows.GroupBy(x => x.NetName + x.PinName + x.Refdes + x.PinNumber).Select(y => y.First()).ToList();

            //Excluding by NetName
            //var netNames = new List<string>();
            //foreach (var filterPin in filterPinDic)
            //    netNames.AddRange(filterPin.Value);
            //rows = rows.Where(x => netNames.Contains(x.NetName)).ToList();

            //rows = rows.Where(x => !x.NetName.Equals("Gnd", StringComparison.CurrentCultureIgnoreCase)).ToList();
            return rows;
        }

        public Dictionary<string, List<string>> GetFilterPins(PinFilterSheet pinFilterSheet)
        {
            if (pinFilterSheet == null)
                return null;

            List<string> pinList = Rows.Where(x => Regex.IsMatch(x.NetName, @"^S\d", RegexOptions.IgnoreCase))
                .Select(x => x.NetName).Distinct().ToList();
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
            List<IGrouping<string, PinFilterRow>> groups = pinFilterSheet.Rows.GroupBy(x => x.Field).ToList();
            foreach (IGrouping<string, PinFilterRow> group in groups)
            {
                List<string> oneField = new List<string>();
                foreach (string data in pinList)
                {
                    if (data == "S0TO7_MAIN_BUFFER1_OUTPUT")
                    {
                    }
                    string pinName = RelayItem.GetPinName(data);
                    foreach (PinFilterRow row in group)
                    {
                        if (CheckEquals(row, pinName) &&
                            CheckContains(row, pinName) &&
                            CheckStartsWith(row, pinName) &&
                            CheckEndsWith(row, pinName) &&
                            CheckNotEquals(row, pinName) &&
                            CheckNotContains(row, pinName) &&
                            CheckNotStartsWith(row, pinName) &&
                            CheckNotEndsWith(row, pinName))
                        {
                            if (!oneField.Contains(data))
                            { oneField.Add(data); }
                        }
                    }
                }

                dic.Add(group.First().Field, oneField);
            }

            return dic;
        }

        public List<string> GetResourcePins(Dictionary<string, List<string>> filterPinDic)
        {
            if (filterPinDic.ContainsKey("Resource Pin"))
            {
                return filterPinDic["Resource Pin"].Select(x => RelayItem.GetPinName(x)).Distinct().ToList();
            }

            return null;
        }

        public List<string> GetDevicePins(Dictionary<string, List<string>> filterPinDic)
        {
            if (filterPinDic.ContainsKey("Device Pin"))
            {
                return filterPinDic["Device Pin"].Select(x => RelayItem.GetPinName(x)).Distinct().ToList();
            }

            return null;
        }

        public List<AdgMatrix> GetAdgMatrix(List<string> sequences, List<ComPinRow> comPinRows)
        {
            List<string> resourcePinName = new List<string>
            {
                "S1",
                "S2",
                "S3",
                "S4",
                "S5",
                "S6",
                "S7",
                "S8"
            };
            List<string> devicePinName = new List<string>
            {
                "D1",
                "D2",
                "D3",
                "D4",
                "D5",
                "D6",
                "D7",
                "D8"
            };
            List<AdgMatrix> adgMatrixList = new List<AdgMatrix>();
            foreach (string sequence in sequences)
            {
                AdgMatrix adgMatrix = new AdgMatrix
                {
                    Name = sequence,
                    ResourcePins = comPinRows.Where(x =>
                        x.Refdes.Equals(sequence, StringComparison.CurrentCultureIgnoreCase) &&
                        resourcePinName.Contains(x.PinName.ToUpper())).ToList(),
                    DevicePins = comPinRows.Where(x =>
                        x.Refdes.Equals(sequence, StringComparison.CurrentCultureIgnoreCase) &&
                        devicePinName.Contains(x.PinName.ToUpper())).ToList()
                };
                adgMatrixList.Add(adgMatrix);
            }
            return adgMatrixList;
        }

        public List<string> GetAdgMatrixSequence(AdgMatrixSheet adgMatrixSheet)
        {
            Dictionary<string, string> sequence = new Dictionary<string, string>();
            List<ComPinRow> rows = Rows
                .Where(x => adgMatrixSheet.Rows.Select(y => y.AdgMatrix.ToUpper()).Contains(x.Refdes))
                .Where(x => x.PinType.Equals("IN", StringComparison.CurrentCultureIgnoreCase) ||
                            x.PinType.Equals("BI", StringComparison.CurrentCultureIgnoreCase)).Where(x =>
                    x.NetName.StartsWith("S0_ADG1414_", StringComparison.CurrentCultureIgnoreCase)).ToList();

            //Get start point
            foreach (IGrouping<string, ComPinRow> group in rows.GroupBy(x => x.Refdes))
            {
                if (group.Any(x => x.PinType.Equals("IN", StringComparison.CurrentCultureIgnoreCase)) &&
                    !group.Any(x => x.PinType.Equals("BI", StringComparison.CurrentCultureIgnoreCase)))
                {
                    sequence.Add(group.First().Refdes, group.First().NetName);
                    break;
                }
            }

            while (SearchSequence(rows, sequence))
            {
            }

            return sequence.Select(x => x.Key).ToList();
        }

        public bool SearchSequence(List<ComPinRow> comPinRows, Dictionary<string, string> sequence)
        {
            if (comPinRows.Where(x => !sequence.ContainsKey(x.Refdes))
                .Any(x => x.NetName.Equals(sequence.Last().Value)))
            {
                ComPinRow row = comPinRows.Where(x => !sequence.ContainsKey(x.Refdes))
                    .First(x => x.NetName.Equals(sequence.Last().Value));
                string refdes = row.Refdes;
                if (comPinRows.Any(x => x.Refdes.Equals(refdes, StringComparison.CurrentCultureIgnoreCase) &&
                                        x.PinType.Equals("IN", StringComparison.CurrentCultureIgnoreCase)))
                {
                    IEnumerable<ComPinRow> inRow = comPinRows.Where(x =>
                        x.Refdes.Equals(refdes, StringComparison.CurrentCultureIgnoreCase) &&
                        x.PinType.Equals("IN", StringComparison.CurrentCultureIgnoreCase));
                    sequence.Add(row.Refdes, inRow.First().NetName);
                    return true;
                }

                sequence.Add(row.Refdes, "");
                return false;
            }

            return false;
        }

        private bool IsEquals(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => data.Equals(x, StringComparison.CurrentCultureIgnoreCase));
        }

        private bool IsContains(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => data.ToUpper().Contains(x.ToUpper()));
        }

        private bool IsStartsWith(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => data.StartsWith(x, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsEndsWith(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => data.EndsWith(x, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsNotEquals(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => !data.Equals(x, StringComparison.CurrentCultureIgnoreCase));
        }

        private bool IsNotContains(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => !data.ToUpper().Contains(x.ToUpper()));
        }

        private bool IsNotStartsWith(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => !data.StartsWith(x, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsNotEndsWith(string data, string checkOr)
        {
            return checkOr.Split('&').All(x => !data.EndsWith(x, StringComparison.OrdinalIgnoreCase));
        }

        private bool CheckEquals(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.Equal))
            {
                return true;
            }

            if (pinFilterRow.Equal.Contains("|") || pinFilterRow.Equal.Contains("&"))
            {
                return pinFilterRow.Equal.Split('|').Any(checkOr => IsEquals(data, checkOr));
            }

            return data.Equals(pinFilterRow.Equal, StringComparison.CurrentCultureIgnoreCase);
        }

        private bool CheckContains(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.Contain))
            {
                return true;
            }

            if (pinFilterRow.Contain.Contains("|") || pinFilterRow.Contain.Contains("&"))
            {
                return pinFilterRow.Contain.Split('|').Any(checkOr => IsContains(data, checkOr));
            }

            return data.ToUpper().Contains(pinFilterRow.Contain.ToUpper());
        }

        private bool CheckStartsWith(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.Prefixed))
            {
                return true;
            }

            if (pinFilterRow.Prefixed.Contains("|") || pinFilterRow.Prefixed.Contains("&"))
            {
                return pinFilterRow.Prefixed.Split('|').Any(checkOr => IsStartsWith(data, checkOr));
            }

            return data.StartsWith(pinFilterRow.Prefixed, StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckEndsWith(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.Suffixed))
            {
                return true;
            }

            if (pinFilterRow.Suffixed.Contains("|") || pinFilterRow.Suffixed.Contains("&"))
            {
                return pinFilterRow.Suffixed.Split('|').Any(checkOr => IsEndsWith(data, checkOr));
            }

            return data.EndsWith(pinFilterRow.Suffixed, StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckNotEquals(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.NotEqual))
            {
                return true;
            }

            if (pinFilterRow.NotEqual.Contains("|") || pinFilterRow.NotEqual.Contains("&"))
            {
                return pinFilterRow.NotEqual.Split('|').Any(checkOr => IsNotEquals(data, checkOr));
            }

            return !data.Equals(pinFilterRow.NotEqual, StringComparison.CurrentCultureIgnoreCase);
        }

        private bool CheckNotContains(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.NotContain))
            {
                return true;
            }

            if (pinFilterRow.NotContain.Contains("|") || pinFilterRow.NotContain.Contains("&"))
            {
                return pinFilterRow.NotContain.Split('|').Any(checkOr => IsNotContains(data, checkOr));
            }

            return !data.ToUpper().Contains(pinFilterRow.NotContain.ToUpper());
        }

        private bool CheckNotStartsWith(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.NotPrefixed))
            {
                return true;
            }

            if (pinFilterRow.NotPrefixed.Contains("|") || pinFilterRow.NotPrefixed.Contains("&"))
            {
                return pinFilterRow.NotPrefixed.Split('|').Any(checkOr => IsNotStartsWith(data, checkOr));
            }

            return !data.StartsWith(pinFilterRow.NotPrefixed, StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckNotEndsWith(PinFilterRow pinFilterRow, string data)
        {
            if (string.IsNullOrEmpty(pinFilterRow.NotSuffixed))
            {
                return true;
            }

            if (pinFilterRow.NotSuffixed.Contains("|") || pinFilterRow.NotSuffixed.Contains("&"))
            {
                return pinFilterRow.NotSuffixed.Split('|').Any(checkOr => IsNotEndsWith(data, checkOr));
            }

            return !data.EndsWith(pinFilterRow.NotSuffixed, StringComparison.OrdinalIgnoreCase);
        }

        #region Field

        #endregion

        #region Properity

        public string Name { get; set; }
        public List<ComPinRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndexDic = new Dictionary<string, int>();

        #endregion
    }

    public class ComPinReader
    {
        private const string HeaderRefdes = "REFDES";
        private const string HeaderPinNumber = "PIN_NUMBER";
        private const string HeaderCompDeviceType = "COMP_DEVICE_TYPE";
        private const string HeaderPinType = "PIN_TYPE";
        private const string HeaderPinName = "PIN_NAME";
        private const string HeaderNetName = "NET_NAME";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"REFDES", true},
            {"PIN_NUMBER", true},
            {"COMP_DEVICE_TYPE", true},
            {"PIN_TYPE", true},
            {"PIN_NAME", true},
            {"NET_NAME", true}
        };

        private int _compDeviceTypeIndex = -1;
        private ComPinSheet _comPinSheet;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _netNameIndex = -1;
        private int _pinNameIndex = -1;
        private int _pinNumberIndex = -1;
        private int _pinTypeIndex = -1;
        private int _refdesIndex = -1;
        private string _name;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public ComPinSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _comPinSheet = new ComPinSheet(_name);

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

            _comPinSheet = ReadSheetData();

            return _comPinSheet;
        }

        private ComPinSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                ComPinRow row = new ComPinRow(_name) { RowNum = i };
                if (_refdesIndex != -1)
                {
                    row.Refdes = _excelWorksheet.GetMergeCellValue(i, _refdesIndex).Trim();
                }

                if (_pinNumberIndex != -1)
                {
                    int value;
                    int.TryParse(_excelWorksheet.GetMergeCellValue(i, _pinNumberIndex).Trim(), out value);
                    row.PinNumber = value;
                }

                if (_compDeviceTypeIndex != -1)
                {
                    row.CompDeviceType = _excelWorksheet.GetMergeCellValue(i, _compDeviceTypeIndex).Trim();
                }

                if (_pinTypeIndex != -1)
                {
                    row.PinType = _excelWorksheet.GetMergeCellValue(i, _pinTypeIndex).Trim();
                }

                if (_pinNameIndex != -1)
                {
                    row.PinName = _excelWorksheet.GetMergeCellValue(i, _pinNameIndex).Trim();
                }

                if (_netNameIndex != -1)
                {
                    row.NetName = _excelWorksheet.GetMergeCellValue(i, _netNameIndex).Trim();
                }

                if (!string.IsNullOrEmpty(row.Refdes) ||
                    !string.IsNullOrEmpty(row.PinNumber.ToString()) ||
                    !string.IsNullOrEmpty(row.PinName) ||
                    !string.IsNullOrEmpty(row.NetName))
                {
                    _comPinSheet.Rows.Add(row);
                }
            }

            return _comPinSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderRefdes, StringComparison.OrdinalIgnoreCase))
                {
                    _refdesIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderRefdes, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderPinNumber, StringComparison.OrdinalIgnoreCase))
                {
                    _pinNumberIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderPinNumber, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCompDeviceType, StringComparison.OrdinalIgnoreCase))
                {
                    _compDeviceTypeIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderCompDeviceType, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderPinType, StringComparison.OrdinalIgnoreCase))
                {
                    _pinTypeIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderPinType, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderPinName, StringComparison.OrdinalIgnoreCase))
                {
                    _pinNameIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderPinName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNetName, StringComparison.OrdinalIgnoreCase))
                {
                    _netNameIndex = i;
                    _comPinSheet.HeaderIndexDic.Add(HeaderNetName, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _comPinSheet.HeaderIndexDic)
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
                        .Equals(HeaderRefdes, StringComparison.OrdinalIgnoreCase))
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
            _refdesIndex = -1;
            _pinNumberIndex = -1;
            _compDeviceTypeIndex = -1;
            _pinTypeIndex = -1;
            _pinNameIndex = -1;
            _netNameIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> dictionaries = new List<Dictionary<string, string>>();
            foreach (ComPinRow row in _comPinSheet.Rows)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>
                {
                    {"REFDES", row.Refdes},
                    {"PIN_NUMBER", row.PinNumber.ToString()},
                    {"COMP_DEVICE_TYPE", row.CompDeviceType},
                    {"PIN_TYPE", row.PinType},
                    {"PIN_NAME", row.PinName},
                    {"NET_NAME", row.NetName}
                };
                dictionaries.Add(dic);
            }

            return dictionaries;
        }
    }
}