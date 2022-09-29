using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local.Const;
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
        #region Constructor

        public PmicLeakageSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicLeakageRow>();
        }

        #endregion

        public InstanceSheet GenInsSheet(string sheetName)
        {
            var instanceSheet = new InstanceSheet(sheetName);
            instanceSheet.AddHeaderFooter(PmicConst.DcLeakage);
            var instanceNames = GetLegalInstanceNames();
            foreach (var instanceName in instanceNames)
                instanceSheet.InstanceRows.Add(GenInstanceRow(instanceName));
            return instanceSheet;
        }

        public List<string> GetLegalInstanceNames()
        {
            var instanceNames = new List<string>();
            Rows.ForEach(row =>
            {
                var instanceName = _func(row);
                if (!string.IsNullOrEmpty(instanceName)) instanceNames.Add(instanceName);
            });
            return instanceNames.Distinct().ToList();
        }

        public List<PmicLeakageRow> GetLegalInstanceNameRows()
        {
            var instanceNameRows = new List<PmicLeakageRow>();
            Rows.ForEach(row =>
            {
                var instanceName = _func(row);
                if (!string.IsNullOrEmpty(instanceName)) instanceNameRows.Add(row);
            });
            return instanceNameRows;
        }

        public List<PmicLeakageRow> GetInlegalInstanceNameRows()
        {
            var legalInstancesNames = GetLegalInstanceNameRows();
            return Rows.Except(legalInstancesNames).ToList();
        }

        public Dictionary<string, List<Tuple<string, PmicLeakageRow>>> GetLegalInstanceNameAndTimeSet()
        {
            var dic = new Dictionary<string, List<Tuple<string, PmicLeakageRow>>>();
            var leakageRows = GetLegalInstanceNameRows();
            foreach (var leakageRow in leakageRows)
                if (dic.ContainsKey(leakageRow.InstanceName))
                {
                    if (!string.IsNullOrEmpty(leakageRow.TimeSetDefine))
                    {
                        var instance = dic[leakageRow.InstanceName].Find(o =>
                            o.Item1.Equals(leakageRow.TimeSetDefine, StringComparison.CurrentCultureIgnoreCase));
                        if (instance == null)
                            dic[leakageRow.InstanceName].Add(Tuple.Create(leakageRow.TimeSetDefine, leakageRow));
                    }
                }
                else
                {
                    dic.Add(leakageRow.InstanceName, new List<Tuple<string, PmicLeakageRow>>());
                    if (!string.IsNullOrEmpty(leakageRow.TimeSetDefine))
                        dic[leakageRow.InstanceName].Add(Tuple.Create(leakageRow.TimeSetDefine, leakageRow));
                }

            return dic;
        }

        private InstanceRow GenInstanceRow(string instanceName)
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
                var timeSetList = instanceNameAndTimeSetDic[instanceName];
                if (timeSetList.Count == 1) instanceRow.TimeSets = timeSetList.First().Item1;
            }

            //instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
            instanceRow.PinLevels = "Levels_Analog";
            return instanceRow;
        }

        public SubFlowSheet GenFlowSheet(string sheetName)
        {
            var subFlowSheet = new SubFlowSheet(sheetName);
            subFlowSheet.FlowRows.AddStartRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);
            subFlowSheet.AddRows(GenFlowBodyRows());
            subFlowSheet.FlowRows.AddEndRows(subFlowSheet.SheetName, SubFlowSheet.Ttime, false);
            return subFlowSheet;
        }

        private FlowRows GenFlowBodyRows()
        {
            var flowRows = new FlowRows();
            var instanceNames = GetLegalInstanceNames();
            foreach (var instanceName in instanceNames)
                flowRows.AddRange(GenFlowRows(instanceName));
            return flowRows;
        }

        private FlowRows GenFlowRows(string instanceName)
        {
            var flowRows = new FlowRows();
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeTest;
            flowRow.Parameter = instanceName;
            flowRow.FailAction = ("F_" + instanceName).AddBlockFlag(BlockBinTableName);
            flowRows.Add(flowRow);

            flowRows.Add_A_Enable_MP_SBIN(BlockBinTableName);

            var bintableRow = new FlowRow();
            bintableRow.OpCode = FlowRow.OpCodeBinTable;
            bintableRow.Parameter = "Bin_" + instanceName;
            flowRows.Add(bintableRow);
            return flowRows;
        }

        public BinTableRows GenBinTableRows()
        {
            var binTableRows = new BinTableRows();
            binTableRows.GenBlockBinTable(BlockBinTableName);
            var instanceNames = GetLegalInstanceNames();
            foreach (var instanceName in instanceNames)
                binTableRows.Add(GenBinTableRow(instanceName));
            binTableRows.GenSetError(PmicConst.DcLeakage);
            return binTableRows;
        }

        private BinTableRow GenBinTableRow(string instanceName)
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
            return binTableRow;
        }

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

        private readonly Func<PmicLeakageRow, string> _func = row =>
        {
            //string matchPatternString = @"(\w*_(Low|High))(\([A-Za-z0-9_]+\))?$";
            var matchPatternString = @"(\w*_(Low|High))$";
            var match = Regex.Match(row.InstanceName, matchPatternString, RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].ToString().Trim();
            return string.Empty;
        };

        private string BlockBinTableName
        {
            get { return PmicConst.Leakage; }
        }

        #endregion
    }

    public class PmicLeakageReader
    {
        private const string ConHeaderInstanceName = "Instance Name";
        private const string ConHeaderPatternName = "Pattern Name";
        private const string ConHeaderMeasurePin = "Measure Pin";
        private const string ConHeaderForceV = "Force V";
        private const string ConHeaderWaitTime = "Wait Time";
        private const string ConHeaderCpftqasnv = "CP_FT_QAS_NV";
        private const string ConHeaderQacnv = "QAC_NV";
        private const string ConHeaderCharnv = "CHAR_NV";
        private const string ConHeaderCharlv = "CHAR_LV";
        private const string ConHeaderCharhv = "CHAR_HV";
        private const string ConHeaderCharulv = "CHAR_ULV";
        private const string ConHeaderRelayOn = "Relay_On";
        private const string ConHeaderRelayOff = "Relay_Off";
        private const string ConHeaderTimeSetDefine = "TimeSet Define(Optional)";
        private int _charHvIndex = -1;
        private int _charLvIndex = -1;
        private int _charNvIndex = -1;
        private int _charUlvIndex = -1;
        private int _cpFtQasNvIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _forceVIndex = -1;
        private int _instanceNameIndex = -1;
        private int _measurePinIndex = -1;
        private int _patternNameIndex = -1;
        private PmicLeakageSheet _pmicLeakageSheet;
        private int _qacNvIndex = -1;
        private int _relayOffIndex = -1;
        private int _relayOnIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _timeSetDefineIndex = -1;
        private int _waitTimeIndex = -1;

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
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PmicLeakageRow(_sheetName);
                row.RowNum = i;
                if (_instanceNameIndex != -1)
                    row.InstanceName = _excelWorksheet.GetMergedCellValue(i, _instanceNameIndex)
                        .Trim();
                if (_patternNameIndex != -1)
                    row.PatternName = _excelWorksheet.GetMergedCellValue(i, _patternNameIndex).Trim();
                if (_measurePinIndex != -1)
                    row.MeasurePin = _excelWorksheet.GetMergedCellValue(i, _measurePinIndex).Trim();
                if (_forceVIndex != -1)
                    row.ForceV = _excelWorksheet.GetMergedCellValue(i, _forceVIndex).Trim();
                if (_waitTimeIndex != -1)
                    row.WaitTime = _excelWorksheet.GetMergedCellValue(i, _waitTimeIndex).Trim();
                if (_cpFtQasNvIndex != -1)
                    row.CpFtQasNv = _excelWorksheet.GetMergedCellValue(i, _cpFtQasNvIndex).Trim();
                if (_qacNvIndex != -1)
                    row.QacNv = _excelWorksheet.GetMergedCellValue(i, _qacNvIndex).Trim();
                if (_charNvIndex != -1)
                    row.CharNv = _excelWorksheet.GetMergedCellValue(i, _charNvIndex).Trim();
                if (_charLvIndex != -1)
                    row.CharLv = _excelWorksheet.GetMergedCellValue(i, _charLvIndex).Trim();
                if (_charHvIndex != -1)
                    row.CharHv = _excelWorksheet.GetMergedCellValue(i, _charHvIndex).Trim();
                if (_charUlvIndex != -1)
                    row.CharUlv = _excelWorksheet.GetMergedCellValue(i, _charUlvIndex).Trim();
                if (_relayOnIndex != -1)
                    row.RelayOn = _excelWorksheet.GetMergedCellValue(i, _relayOnIndex).Trim();
                if (_relayOffIndex != -1)
                    row.RelayOff = _excelWorksheet.GetMergedCellValue(i, _relayOffIndex).Trim();
                if (_timeSetDefineIndex != -1)
                    row.TimeSetDefine = _excelWorksheet.GetMergedCellValue(i, _timeSetDefineIndex)
                        .Trim();
                if (!string.IsNullOrEmpty(row.InstanceName) && !string.IsNullOrEmpty(row.MeasurePin))
                    _pmicLeakageSheet.Rows.Add(row);
            }

            _pmicLeakageSheet.InstanceNameIndex = _instanceNameIndex;
            _pmicLeakageSheet.PatternNameIndex = _patternNameIndex;
            _pmicLeakageSheet.MeasurePinIndex = _measurePinIndex;
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
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
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
                    _measurePinIndex = i;
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

                if (lStrHeader.Equals(ConHeaderCpftqasnv, StringComparison.OrdinalIgnoreCase))
                {
                    _cpFtQasNvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderQacnv, StringComparison.OrdinalIgnoreCase))
                {
                    _qacNvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCharnv, StringComparison.OrdinalIgnoreCase))
                {
                    _charNvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCharlv, StringComparison.OrdinalIgnoreCase))
                {
                    _charLvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCharhv, StringComparison.OrdinalIgnoreCase))
                {
                    _charHvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCharulv, StringComparison.OrdinalIgnoreCase))
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
                    _timeSetDefineIndex = i;
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
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
            _measurePinIndex = -1;
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