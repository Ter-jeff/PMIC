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
    public class PmicIdsRow
    {
        #region Property

        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string InstanceName { set; get; }
        public string MeasurePin { set; get; }
        public string CpFtQasNv { set; get; }
        public string QacNv { set; get; }
        public string CharNv { set; get; }
        public string CharLv { set; get; }
        public string CharHv { set; get; }
        public string CharUlv { set; get; }

        #endregion

        #region Constructor

        public PmicIdsRow()
        {
        }

        public PmicIdsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }

        #endregion
    }

    public class PmicIdsSheet
    {
        #region Constructor

        public PmicIdsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicIdsRow>();
        }

        #endregion

        public InstanceSheet GenInsSheet(string sheetName)
        {
            var instanceSheet = new InstanceSheet(sheetName);
            instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_NV");
            instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_LV");
            instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_HV");

            var instanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            foreach (var instanceName in instanceNames)
                instanceSheet.InstanceRows.AddRange(GenInstanceRows(instanceName, "NV"));
            return instanceSheet;
        }

        private List<InstanceRow> GenInstanceRows(string instanceName, string voltage)
        {
            var instanceRows = new List<InstanceRow>();
            var dcSelector = GetDcSelector(voltage);
            var instanceRow = new InstanceRow();
            if (!instanceName.ToLower().Contains("pretrim"))
            {
                instanceRow.Type = "VBT";
                instanceRow.TestName = instanceName + "_PreSetUp";
                instanceRow.Name = "DC_IDS_UVI80_PreSetup";
                instanceRow.DcCategory = "Analog";
                instanceRow.DcSelector = dcSelector;
                instanceRow.AcCategory = "Common";
                instanceRow.AcSelector = "Typ";
                instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
                instanceRow.PinLevels = "Levels_Analog";
                instanceRows.Add(instanceRow);
            }

            instanceRow = new InstanceRow();
            instanceRow.Type = "VBT";
            instanceRow.TestName = instanceName;
            instanceRow.Name = "DC_IDS_UVI80";
            instanceRow.DcCategory = "Analog";
            instanceRow.DcSelector = dcSelector;
            instanceRow.AcCategory = "Common";
            instanceRow.AcSelector = "Typ";
            instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
            instanceRow.PinLevels = "Levels_Analog";
            instanceRows.Add(instanceRow);
            return instanceRows;
        }

        private string GetDcSelector(string voltage)
        {
            var dcSelector = "Typ";
            if (voltage.Equals("HV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Max";
            if (voltage.Equals("LV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Min";
            return dcSelector;
        }

        public List<SubFlowSheet> GenSubFlowSheets(string sheetName)
        {
            var subFlowSheets = new List<SubFlowSheet>();
            var voltages = new List<string> { "NV", "HV", "LV" };
            foreach (var voltage in voltages)
            {
                var flowSheetName = sheetName + "_" + voltage;
                subFlowSheets.Add(GenFlowSheet(flowSheetName, voltage));
            }

            return subFlowSheets;
        }

        public SubFlowSheet GenFlowSheet(string flowSheetName, string voltage)
        {
            var subFlowSheet = new SubFlowSheet(flowSheetName);
            subFlowSheet.FlowRows.AddStartRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);
            subFlowSheet.AddRows(GenFlowBodyRows(voltage));
            subFlowSheet.FlowRows.AddEndRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);
            return subFlowSheet;
        }

        private FlowRows GenFlowBodyRows(string voltage)
        {
            var flowRows = new FlowRows();
            var instanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            foreach (var instanceName in instanceNames)
            {
                if (instanceName.ToLower().Contains("pretrim")) continue;
                flowRows.AddRange(GenFlowRows(instanceName, voltage));
            }

            return flowRows;
        }

        private FlowRows GenFlowRows(string instanceName, string voltage)
        {
            var flowRows = new FlowRows();
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeTest;
            flowRow.Parameter = instanceName + "_PreSetUp";
            flowRow.FailAction = ("F_" + instanceName + "_" + voltage).AddBlockFlag(Block);
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeTest;
            flowRow.Parameter = instanceName;
            flowRow.FailAction = ("F_" + instanceName + "_" + voltage).AddBlockFlag(Block);
            flowRows.Add(flowRow);

            flowRows.Add_A_Enable_MP_SBIN(PmicConst.Ids);

            flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = "Bin_" + instanceName + "_" + voltage;
            flowRows.Add(flowRow);
            return flowRows;
        }

        public BinTableRows GenBinTableRows()
        {
            var binTableRows = new BinTableRows();
            binTableRows.GenBlockBinTable(PmicConst.Ids);
            var instanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            var voltages = new List<string> { "NV", "HV", "LV" };
            foreach (var voltage in voltages)
                foreach (var instanceName in instanceNames)
                {
                    if (instanceName.ToLower().Contains("pretrim"))
                        continue;

                    binTableRows.Add(GenBinTableRows(instanceName, voltage));
                }

            binTableRows.GenSetError(PmicConst.Ids);
            return binTableRows;
        }

        private BinTableRow GenBinTableRows(string instanceName, string type)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.Pmic, SheetName);
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = "Bin_" + instanceName + "_" + type;
            binTableRow.ItemList = "F_" + instanceName + "_" + type;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            return binTableRow;
        }

        public List<string> GetMeasurePins()
        {
            return Rows.Select(o => o.MeasurePin.Trim()).Distinct().ToList();
        }

        #region Field

        internal int InstanceNameIndex;
        internal int MeasurePinIndex;
        internal int CpFtQasNvIndex;
        internal int QacNvIndex;
        internal int CharNvIndex;
        internal int CharLvIndex;
        internal int CharHvIndex;
        internal int CharUlvIndex;

        #endregion

        #region Property

        public string SheetName { get; set; }
        public string Block
        {
            get { return Regex.Replace(SheetName, "^PMIC_", "", RegexOptions.IgnoreCase); }
        }

        public List<PmicIdsRow> Rows { get; set; }

        #endregion
    }

    public class PmicIdsReader
    {
        private const string ConHeaderInstanceName = "Instance Name";
        private const string ConHeaderMeasurePin = "Measure Pin";
        private const string ConHeaderCpftqasnv = "CP_FT_QAS_NV";
        private const string ConHeaderQacnv = "QAC_NV";
        private const string ConHeaderCharnv = "CHAR_NV";
        private const string ConHeaderCharlv = "CHAR_LV";
        private const string ConHeaderCharhv = "CHAR_HV";
        private const string ConHeaderCharulv = "CHAR_ULV";
        private int _charHvIndex = -1;
        private int _charLvIndex = -1;
        private int _charNvIndex = -1;
        private int _charUlvIndex = -1;
        private int _cpFtQasNvIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _instanceNameIndex = -1;
        private int _measurePinIndex = -1;
        private PmicIdsSheet _pmicIdsSheet;
        private int _qacNvIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public PmicIdsSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _pmicIdsSheet = new PmicIdsSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _pmicIdsSheet = ReadSheetData();

            return _pmicIdsSheet;
        }

        private PmicIdsSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PmicIdsRow(_sheetName);
                row.RowNum = i;
                if (_instanceNameIndex != -1)
                    row.InstanceName = _excelWorksheet.GetMergedCellValue(i, _instanceNameIndex)
                        .Trim();
                if (_measurePinIndex != -1)
                    row.MeasurePin = _excelWorksheet.GetMergedCellValue(i, _measurePinIndex).Trim();
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
                if (!string.IsNullOrEmpty(row.InstanceName) && !string.IsNullOrEmpty(row.MeasurePin))
                    _pmicIdsSheet.Rows.Add(row);
            }

            _pmicIdsSheet.InstanceNameIndex = _instanceNameIndex;
            _pmicIdsSheet.MeasurePinIndex = _measurePinIndex;
            _pmicIdsSheet.CpFtQasNvIndex = _cpFtQasNvIndex;
            _pmicIdsSheet.QacNvIndex = _qacNvIndex;
            _pmicIdsSheet.CharNvIndex = _charNvIndex;
            _pmicIdsSheet.CharLvIndex = _charLvIndex;
            _pmicIdsSheet.CharHvIndex = _charHvIndex;
            _pmicIdsSheet.CharUlvIndex = _charUlvIndex;

            return _pmicIdsSheet;
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

                if (lStrHeader.Equals(ConHeaderMeasurePin, StringComparison.OrdinalIgnoreCase))
                {
                    _measurePinIndex = i;
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

                if (lStrHeader.Equals(ConHeaderCharulv, StringComparison.OrdinalIgnoreCase)) _charUlvIndex = i;
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
            _measurePinIndex = -1;
            _cpFtQasNvIndex = -1;
            _qacNvIndex = -1;
            _charNvIndex = -1;
            _charLvIndex = -1;
            _charHvIndex = -1;
            _charUlvIndex = -1;
        }
    }
}