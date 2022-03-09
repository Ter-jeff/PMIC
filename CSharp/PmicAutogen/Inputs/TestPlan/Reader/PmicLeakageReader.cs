using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader
{

    public class PmicLeakageRow
    {
        #region Property
        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string InstanceName { set; get; }
        public string PatternName { set; get; }
        public string MeasurePin { set; get; }
        public string ForceV { set; get; }
        public string WaitTime { set; get; }
        public string CpFtQasNv { set; get; }
        public string QacNv { set; get; }
        public string CharNv { set; get; }
        public string CharLv { set; get; }
        public string CharHv { set; get; }
        public string CharUlv { set; get; }
        public string RelayOn { set; get; }
        public string RelayOff { set; get; }
        public string TimeSetDefine { set; get; }
        #endregion

        #region Constructor
        public PmicLeakageRow()
        {
        }

        public PmicLeakageRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }
        #endregion
    }

    public class PmicLeakageSheet
    {
        #region Field
        internal int InstanceNameIndex;
        internal int PatternNameIndex;
        internal int MeasurePinIndex;
        internal int ForceVIndex;
        internal int WaitTimeIndex;
        internal int CpFtQasNvIndex;
        internal int QacNvIndex;
        internal int CharNvIndex;
        internal int CharLvIndex;
        internal int CharHvIndex;
        internal int CharUlvIndex;
        internal int RelayOnIndex;
        internal int RelayOffIndex;
        internal int TimeSetDefineIndex;
        #endregion

        #region Property
        public string SheetName { get; set; }
        public List<PmicLeakageRow> Rows { get; set; }

        private Func<PmicLeakageRow, string> func = row =>
        {
            //string matchPatternString = @"(\w*_(Low|High))(\([A-Za-z0-9_]+\))?$";
            string matchPatternString = @"(\w*_(Low|High))$";
            Match match = Regex.Match(row.InstanceName, matchPatternString, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].ToString().Trim();
            }
            return string.Empty;
        };
        #endregion

        #region Constructor
        public PmicLeakageSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicLeakageRow>();
        }

        #endregion
        public List<InstanceRow> GenInstance()
        {
            var instanceRows = new List<InstanceRow>();
            var InstanceNames = GetLegalInstanceNames();
            foreach (var instanceName in InstanceNames)
                GenInstanceRows(instanceName, instanceRows);
            return instanceRows;
        }

        public List<string> GetLegalInstanceNames()
        {
            var InstanceNames = new List<string>();
            Rows.ForEach(row => {
                string instanceName = func(row);
                if (!string.IsNullOrEmpty(instanceName))
                {
                    InstanceNames.Add(instanceName);
                }
            });
            return InstanceNames.Distinct().ToList();
        }

        public List<PmicLeakageRow> GetLegalInstanceNameRows()
        {
            var InstanceNameRows = new List<PmicLeakageRow>();
            Rows.ForEach(row => {
                string instanceName = func(row);
                if (!string.IsNullOrEmpty(instanceName))
                {
                    InstanceNameRows.Add(row);
                }
            });
            return InstanceNameRows;
        }

        public List<PmicLeakageRow> GetInlegalInstanceNameRows()
        {
            var legalInstancesNames = GetLegalInstanceNameRows();
            return Rows.Except(legalInstancesNames).ToList();
        }


        public Dictionary<string, List<Tuple<string, PmicLeakageRow>>> GetLegalInstanceNameAndTimeSet()
        {
            Dictionary<string, List<Tuple<string, PmicLeakageRow>>> dic = new Dictionary<string, List<Tuple<string, PmicLeakageRow>>>();
            List<PmicLeakageRow> leakageRows = GetLegalInstanceNameRows();
            foreach (var leakageRow in leakageRows)
            {
                if (dic.ContainsKey(leakageRow.InstanceName))
                {
                    if (!string.IsNullOrEmpty(leakageRow.TimeSetDefine))
                    {
                        var instance = dic[leakageRow.InstanceName].Find(o => o.Item1.Equals(leakageRow.TimeSetDefine, StringComparison.CurrentCultureIgnoreCase));
                        if(instance==null)
                            dic[leakageRow.InstanceName].Add(Tuple.Create<string, PmicLeakageRow>(leakageRow.TimeSetDefine, leakageRow));
                    }
                }
                else
                {
                    dic.Add(leakageRow.InstanceName, new List<Tuple<string, PmicLeakageRow>>());
                    if (!string.IsNullOrEmpty(leakageRow.TimeSetDefine))
                        dic[leakageRow.InstanceName].Add(Tuple.Create<string, PmicLeakageRow>(leakageRow.TimeSetDefine, leakageRow));
                }
            }
            return dic;
        }

        private void GenInstanceRows(string instanceName, List<InstanceRow> instanceRows)
        {
            var instanceRow = new InstanceRow();
            instanceRow.Type = "VBT";
            instanceRow.TestName = instanceName;
            instanceRow.Name = "DC_Leak_Test";
            instanceRow.DcCategory = "Leakage";
            instanceRow.DcSelector = "Typ";
            instanceRow.AcCategory = "Common";
            instanceRow.AcSelector = "Typ";
            var instanceNameAndTimeSetDic = GetLegalInstanceNameAndTimeSet();
            if (instanceNameAndTimeSetDic.ContainsKey(instanceName))
            {
                List<Tuple<string, PmicLeakageRow>> timeSetList = instanceNameAndTimeSetDic[instanceName];
                if (timeSetList.Count == 1)
                {
                    instanceRow.TimeSets = timeSetList.First().Item1;
                }
            }
            //instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
            instanceRow.PinLevels = "Levels_Analog";
            instanceRows.Add(instanceRow);

        }

        public SubFlowSheet GenSubFlowSheet(string flowSheetName)
        {
            var flowSheet = new SubFlowSheet(flowSheetName);
            flowSheet.AddStartRows(SubFlowSheet.Ttime);
            flowSheet.AddRows(GenFlowBodyRows());
            flowSheet.AddEndRows(SubFlowSheet.Ttime,false);
            return flowSheet;
        }

        private List<FlowRow> GenFlowBodyRows()
        {
            var flowRows = new List<FlowRow>();
            var InstanceNames = GetLegalInstanceNames();
            foreach (var instanceName in InstanceNames)
                GenFlowRows(instanceName, flowRows);
            return flowRows;
        }

        private void GenFlowRows(string instanceName, List<FlowRow> flowRows)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = "Test";
            flowRow.Parameter = instanceName;
            flowRow.FailAction = "F_" + instanceName;
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = "Bin_" + instanceName;
            flowRows.Add(flowRow);

        }

        public List<BinTableRow> GenBinTableRows()
        {
            var binTableRows = new List<BinTableRow>();
            var InstanceNames = GetLegalInstanceNames();
            foreach (var instanceName in InstanceNames)
                GenBinTableRows(instanceName, binTableRows);
            return binTableRows;
        }

        private void GenBinTableRows(string instanceName, List<BinTableRow> binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.Pmic, SheetName);
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = "Bin_" + instanceName;
            binTableRow.ItemList = "F_" + instanceName;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

    }

    public class PmicLeakageReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private PmicLeakageSheet _pmicLeakageSheet;

        private const string ConHeaderInstanceName = "Instance Name";
        private const string ConHeaderPatternName = "Pattern Name";
        private const string ConHeaderMeasurePin = "Measure Pin";
        private const string ConHeaderForceV = "Force V";
        private const string ConHeaderWaitTime = "Wait Time";
        private const string ConHeaderCPFTQASNV = "CP_FT_QAS_NV";
        private const string ConHeaderQACNV = "QAC_NV";
        private const string ConHeaderCHARNV = "CHAR_NV";
        private const string ConHeaderCHARLV = "CHAR_LV";
        private const string ConHeaderCHARHV = "CHAR_HV";
        private const string ConHeaderCHARULV = "CHAR_ULV";
        private const string ConHeaderRelayOn = "Relay_On";
        private const string ConHeaderRelayOff = "Relay_Off";
        private const string ConHeaderTimeSetDefine = "TimeSet Define(Optional)";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _instanceNameIndex = -1;
        private int _patternNameIndex = -1;
        private int _MeasurePinIndex = -1;
        private int _forceVIndex = -1;
        private int _waitTimeIndex = -1;
        private int _cpFtQasNvIndex = -1;
        private int _qacNvIndex = -1;
        private int _charNvIndex = -1;
        private int _charLvIndex = -1;
        private int _charHvIndex = -1;
        private int _charUlvIndex = -1;
        private int _relayOnIndex = -1;
        private int _relayOffIndex = -1;
        private int _timeSetDefineIndex = -1;

        public PmicLeakageSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _pmicLeakageSheet = new PmicLeakageSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _pmicLeakageSheet = ReadSheetData();

            return _pmicLeakageSheet;
        }

        private PmicLeakageSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PmicLeakageRow row = new PmicLeakageRow(_sheetName);
                row.RowNum = i;
                if (_instanceNameIndex != -1)
                    row.InstanceName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _instanceNameIndex).Trim();
                if (_patternNameIndex != -1)
                    row.PatternName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _patternNameIndex).Trim();
                if (_MeasurePinIndex != -1)
                    row.MeasurePin = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _MeasurePinIndex).Trim();
                if (_forceVIndex != -1)
                    row.ForceV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _forceVIndex).Trim();
                if (_waitTimeIndex != -1)
                    row.WaitTime = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _waitTimeIndex).Trim();
                if (_cpFtQasNvIndex != -1)
                    row.CpFtQasNv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cpFtQasNvIndex).Trim();
                if (_qacNvIndex != -1)
                    row.QacNv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _qacNvIndex).Trim();
                if (_charNvIndex != -1)
                    row.CharNv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _charNvIndex).Trim();
                if (_charLvIndex != -1)
                    row.CharLv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _charLvIndex).Trim();
                if (_charHvIndex != -1)
                    row.CharHv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _charHvIndex).Trim();
                if (_charUlvIndex != -1)
                    row.CharUlv = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _charUlvIndex).Trim();
                if (_relayOnIndex != -1)
                    row.RelayOn = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _relayOnIndex).Trim();
                if (_relayOffIndex != -1)
                    row.RelayOff = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _relayOffIndex).Trim();
                if (_timeSetDefineIndex != -1)
                    row.TimeSetDefine = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _timeSetDefineIndex).Trim();
                if (!string.IsNullOrEmpty(row.InstanceName) && !string.IsNullOrEmpty(row.MeasurePin))
                    _pmicLeakageSheet.Rows.Add(row);
            }

            _pmicLeakageSheet.InstanceNameIndex = _instanceNameIndex;
            _pmicLeakageSheet.PatternNameIndex = _patternNameIndex;
            _pmicLeakageSheet.MeasurePinIndex = _MeasurePinIndex;
            _pmicLeakageSheet.ForceVIndex = _forceVIndex;
            _pmicLeakageSheet.WaitTimeIndex = _waitTimeIndex;
            _pmicLeakageSheet.CpFtQasNvIndex = _cpFtQasNvIndex;
            _pmicLeakageSheet.QacNvIndex = _qacNvIndex;
            _pmicLeakageSheet.CharNvIndex = _charNvIndex;
            _pmicLeakageSheet.CharLvIndex = _charLvIndex;
            _pmicLeakageSheet.CharHvIndex = _charHvIndex;
            _pmicLeakageSheet.CharUlvIndex = _charUlvIndex;
            _pmicLeakageSheet.RelayOnIndex = _relayOnIndex;
            _pmicLeakageSheet.RelayOffIndex = _relayOffIndex;
            _pmicLeakageSheet.TimeSetDefineIndex = _timeSetDefineIndex;

            return _pmicLeakageSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
                {
                    _instanceNameIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderPatternName, StringComparison.OrdinalIgnoreCase))
                {
                    _patternNameIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderMeasurePin, StringComparison.OrdinalIgnoreCase))
                {
                    _MeasurePinIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderForceV, StringComparison.OrdinalIgnoreCase))
                {
                    _forceVIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderWaitTime, StringComparison.OrdinalIgnoreCase))
                {
                    _waitTimeIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCPFTQASNV, StringComparison.OrdinalIgnoreCase))
                {
                    _cpFtQasNvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderQACNV, StringComparison.OrdinalIgnoreCase))
                {
                    _qacNvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARNV, StringComparison.OrdinalIgnoreCase))
                {
                    _charNvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARLV, StringComparison.OrdinalIgnoreCase))
                {
                    _charLvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARHV, StringComparison.OrdinalIgnoreCase))
                {
                    _charHvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderCHARULV, StringComparison.OrdinalIgnoreCase))
                {
                    _charUlvIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderRelayOn, StringComparison.OrdinalIgnoreCase))
                {
                    _relayOnIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderRelayOff, StringComparison.OrdinalIgnoreCase))
                {
                    _relayOffIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTimeSetDefine, StringComparison.OrdinalIgnoreCase))
                {
                    _timeSetDefineIndex = i;
                    continue;
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
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
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
            _instanceNameIndex = -1;
            _patternNameIndex = -1;
            _MeasurePinIndex = -1;
            _forceVIndex = -1;
            _waitTimeIndex = -1;
            _cpFtQasNvIndex = -1;
            _qacNvIndex = -1;
            _charNvIndex = -1;
            _charLvIndex = -1;
            _charHvIndex = -1;
            _charUlvIndex = -1;
            _relayOnIndex = -1;
            _relayOffIndex = -1;
            _timeSetDefineIndex = -1;
        }
    }
}
