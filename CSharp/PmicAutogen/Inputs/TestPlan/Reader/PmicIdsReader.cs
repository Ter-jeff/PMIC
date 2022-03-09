using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using System;
using System.Collections.Generic;
using System.Linq;

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
        public List<PmicIdsRow> Rows { get; set; }
        #endregion

        #region Constructor
        public PmicIdsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicIdsRow>();
        }
        #endregion

        public List<InstanceRow> GenInstance()
        {
            var instanceRows = new List<InstanceRow>();
            var InstanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            foreach (var instanceName in InstanceNames)
                GenInstanceRows(instanceName, "NV", instanceRows);
            //foreach (var row in Rows)
            //    GenInstanceRows(row, "LV", instanceRows);
            //foreach (var row in Rows)
            //    GenInstanceRows(row, "HV", instanceRows);
            return instanceRows;
        }

        private static void GenInstanceRows(string instanceName, string type, List<InstanceRow> instanceRows)
        {
            var dcSelector = "Typ";
            if (type.Equals("HV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Max";
            if (type.Equals("LV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Min";
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
        }

        public List<SubFlowSheet> GenSubFlowSheets(string idsSheet)
        {
            var subFlowSheets = new List<SubFlowSheet>();
            var flowSheetName = "Flow_" + idsSheet + "_NV";
            subFlowSheets.Add(GenFlow(flowSheetName, "NV"));
            flowSheetName = "Flow_" + idsSheet + "_LV";
            subFlowSheets.Add(GenFlow(flowSheetName, "LV"));
            flowSheetName = "Flow_" + idsSheet + "_HV";
            subFlowSheets.Add(GenFlow(flowSheetName, "HV"));
            return subFlowSheets;
        }

        public SubFlowSheet GenFlow(string flowSheetName, string type)
        {
            var flowSheet = new SubFlowSheet(flowSheetName);
            flowSheet.AddStartRows(SubFlowSheet.Ttime);
            flowSheet.AddRows(GenFlowBodyRows(type));
            flowSheet.AddEndRows(SubFlowSheet.Ttime);
            return flowSheet;
        }

        private List<FlowRow> GenFlowBodyRows(string type)
        {
            var flowRows = new List<FlowRow>();
            var InstanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            foreach (var instanceName in InstanceNames)
                GenFlowRows(instanceName, type, flowRows);
            return flowRows;
        }

        private static void GenFlowRows(string instanceName, string type, List<FlowRow> flowRows)
        {
            if (instanceName.ToLower().Contains("pretrim")) return;
            var flowRow = new FlowRow();
            flowRow.OpCode = "Test";
            flowRow.Parameter = instanceName + "_PreSetUp";
            flowRow.FailAction = "F_" + instanceName + "_" + type;
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = "Test";
            flowRow.Parameter = instanceName;
            flowRow.FailAction = "F_" + instanceName + "_" + type;
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = "Bin_" + instanceName + "_" + type;
            flowRows.Add(flowRow);
        }

        public List<BinTableRow> GenBinTableRows()
        {
            var binTableRows = new List<BinTableRow>();
            var InstanceNames = Rows.Select(x => x.InstanceName).Distinct().ToList();
            foreach (var instanceName in InstanceNames)
                GenBinTableRows(instanceName, binTableRows, "NV");
            foreach (var instanceName in InstanceNames)
                GenBinTableRows(instanceName, binTableRows, "HV");
            foreach (var instanceName in InstanceNames)
                GenBinTableRows(instanceName, binTableRows, "LV");
            return binTableRows;
        }

        private void GenBinTableRows(string instanceName, List<BinTableRow> binTableRows, string type)
        {
            if (instanceName.ToLower().Contains("pretrim")) return;
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
            binTableRows.Add(binTableRow);
        }

        public List<string> GetMeasurePins()
        {
            return Rows.Select(o => o.MeasurePin.Trim()).Distinct().ToList();
        }
    }

    public class PmicIdsReader
    {
        private ExcelWorksheet _excelWorksheet;
        private string _sheetName;
        private PmicIdsSheet _pmicIdsSheet;

        private const string ConHeaderInstanceName = "Instance Name";
        private const string ConHeaderMeasurePin = "Measure Pin";
        private const string ConHeaderCPFTQASNV = "CP_FT_QAS_NV";
        private const string ConHeaderQACNV = "QAC_NV";
        private const string ConHeaderCHARNV = "CHAR_NV";
        private const string ConHeaderCHARLV = "CHAR_LV";
        private const string ConHeaderCHARHV = "CHAR_HV";
        private const string ConHeaderCHARULV = "CHAR_ULV";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _instanceNameIndex = -1;
        private int _MeasurePinIndex = -1;
        private int _cpFtQasNvIndex = -1;
        private int _qacNvIndex = -1;
        private int _charNvIndex = -1;
        private int _charLvIndex = -1;
        private int _charHvIndex = -1;
        private int _charUlvIndex = -1;

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
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PmicIdsRow row = new PmicIdsRow(_sheetName);
                row.RowNum = i;
                if (_instanceNameIndex != -1)
                    row.InstanceName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _instanceNameIndex).Trim();
                if (_MeasurePinIndex != -1)
                    row.MeasurePin = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _MeasurePinIndex).Trim();
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
                if(!string.IsNullOrEmpty(row.InstanceName) && !string.IsNullOrEmpty(row.MeasurePin))
                    _pmicIdsSheet.Rows.Add(row);
            }

            _pmicIdsSheet.InstanceNameIndex = _instanceNameIndex;
            _pmicIdsSheet.MeasurePinIndex = _MeasurePinIndex;
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
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderInstanceName, StringComparison.OrdinalIgnoreCase))
                {
                    _instanceNameIndex = i;
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderMeasurePin, StringComparison.OrdinalIgnoreCase))
                {
                    _MeasurePinIndex = i;
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
            _MeasurePinIndex = -1;
            _cpFtQasNvIndex = -1;
            _qacNvIndex = -1;
            _charNvIndex = -1;
            _charLvIndex = -1;
            _charHvIndex = -1;
            _charUlvIndex = -1;
        }

    }
}