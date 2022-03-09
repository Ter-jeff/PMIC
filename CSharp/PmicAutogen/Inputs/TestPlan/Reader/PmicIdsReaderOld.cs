using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class PmicIdsSheetRow
    {
        #region Constructor

        public PmicIdsSheetRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            TestItem = "";
            Ttr = "";
            Step = "";
            Description = "";
            Instance = "";
            MeasurePins = "";
            CurrentRange = "";
            Cp1LoLimit = "";
            Cp1HiLimit = "";
            Comment = "";
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string TestItem { get; set; }
        public string Ttr { get; set; }
        public string Step { get; set; }
        public string Description { get; set; }
        public string Instance { get; set; }
        public string MeasurePins { get; set; }
        public string CurrentRange { get; set; }
        public string Cp1LoLimit { get; set; }
        public string Cp1HiLimit { get; set; }
        public string Comment { get; set; }

        #endregion
    }

    public class PmicIdsSheetSheet
    {
        #region Constructor

        public PmicIdsSheetSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PmicIdsSheetRow>();
        }

        #endregion

        public List<InstanceRow> GenInstance()
        {
            var instanceRows = new List<InstanceRow>();
            foreach (var row in Rows)
                GenInstanceRows(row, "NV", instanceRows);
            foreach (var row in Rows)
                GenInstanceRows(row, "LV", instanceRows);
            foreach (var row in Rows)
                GenInstanceRows(row, "HV", instanceRows);
            return instanceRows;
        }

        private static void GenInstanceRows(PmicIdsSheetRow row, string type, List<InstanceRow> instanceRows)
        {
            var dcSelector = "Typ";
            if (type.Equals("HV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Max";
            if (type.Equals("LV", StringComparison.CurrentCultureIgnoreCase)) dcSelector = "Min";
            var instanceRow = new InstanceRow();
            instanceRow.TestName = row.Instance + "_PreSetUp" + "_" + type;
            instanceRow.Name = "DC_IDS_UVI80_PreSetup";
            instanceRow.DcCategory = "Analog";
            instanceRow.DcSelector = dcSelector;
            instanceRow.AcCategory = "Common";
            instanceRow.AcSelector = "Typ";
            instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
            instanceRow.PinLevels = "Levels_Analog";
            var vbtFunctionBase = TestProgram.VbtFunctionLib.GetFunctionByName(instanceRow.Name);
            vbtFunctionBase.SetParamValue("Measure_pin",
                "VDD_ANA_S_UVI80,VDD_ANA_UVI80,VDD_BOOST_LDO_UVI80,VDD_BOOST_SNS_UVI80,VDD_BOOST_UVI80,VDD_BUCK0_2_7_11_UVI80,VDD_BUCK1_8_9_UVI80,VDD_BUCK3_14_UVI80,VDD_DIG_S_UVI80,VDD_DIG_UVI80,VDD_HI_INT1_UVI80,VDD_HI_INT2_UVI80,VDD_HI_INT3_UVI80,VDD_HI_INT4_UVI80,VDD_HI_INT5_UVI80,VDD_HI_INT6_UVI80,VDD_LDO19_UVI80,VDD_LDO2_UVI80,VDD_LDO3_14_UVI80,VDD_LDO5_UVI80,VDD_MAIN_DRV_UVI80,VDD_MAIN_LDO_UVI80,VDD_MAIN_SNS_UVI80,VDD_MAIN_SNS_WLED_UVI80,VDD_MAIN_UVI80,VDD_MAIN_WBOOST_UVI80,VDD_MAIN_WIDAC_UVI80,VDD_MAIN1_UVI80,VDD_RTC_ALT_UVI80,VDD_SNS_SPARE_UVI80,VDD_SW1_UVI80,VDD_SW2_UVI80,VDD_SW3_UVI80,VDDIO_1V2_UVI80,VDDIO_BUCK3_UVI80,VBAT_UVI80,IBAT_UVI80");
            vbtFunctionBase.SetParamValue("ForceV",
                "1.5,1.5,3.8,3.8,3.8,3.8,3.8,3.8,1.5,1.5,5,5,5,5,5,5,1.5,3.8,1.2,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,1.8,1.8,1.2,1.2,1.8,3.8,3.8,1.8");
            vbtFunctionBase.SetParamValue("MeasureI_Range", "0.0002");
            instanceRow.ArgList = vbtFunctionBase.Parameters;
            instanceRow.Args = vbtFunctionBase.Args;
            instanceRows.Add(instanceRow);

            instanceRow = new InstanceRow();
            instanceRow.TestName = row.Instance + "_" + type;
            instanceRow.Name = "DC_IDS_UVI80";
            instanceRow.DcCategory = "Analog";
            instanceRow.DcSelector = dcSelector;
            instanceRow.AcCategory = "Common";
            instanceRow.AcSelector = "Typ";
            instanceRow.TimeSets = "TIMESET_PMIC_Dummy";
            instanceRow.PinLevels = "Levels_Analog";
            vbtFunctionBase = TestProgram.VbtFunctionLib.GetFunctionByName(instanceRow.Name);
            vbtFunctionBase.SetParamValue("Measure_pin",
                "VDD_ANA_S_UVI80,VDD_ANA_UVI80,VDD_BOOST_LDO_UVI80,VDD_BOOST_SNS_UVI80,VDD_BOOST_UVI80,VDD_BUCK0_2_7_11_UVI80,VDD_BUCK1_8_9_UVI80,VDD_BUCK3_14_UVI80,VDD_DIG_S_UVI80,VDD_DIG_UVI80,VDD_HI_INT1_UVI80,VDD_HI_INT2_UVI80,VDD_HI_INT3_UVI80,VDD_HI_INT4_UVI80,VDD_HI_INT5_UVI80,VDD_HI_INT6_UVI80,VDD_LDO19_UVI80,VDD_LDO2_UVI80,VDD_LDO3_14_UVI80,VDD_LDO5_UVI80,VDD_MAIN_DRV_UVI80,VDD_MAIN_LDO_UVI80,VDD_MAIN_SNS_UVI80,VDD_MAIN_SNS_WLED_UVI80,VDD_MAIN_UVI80,VDD_MAIN_WBOOST_UVI80,VDD_MAIN_WIDAC_UVI80,VDD_MAIN1_UVI80,VDD_RTC_ALT_UVI80,VDD_SNS_SPARE_UVI80,VDD_SW1_UVI80,VDD_SW2_UVI80,VDD_SW3_UVI80,VDDIO_1V2_UVI80,VDDIO_BUCK3_UVI80,VBAT_UVI80,IBAT_UVI80");
            vbtFunctionBase.SetParamValue("ForceV",
                "1.5,1.5,3.8,3.8,3.8,3.8,3.8,3.8,1.5,1.5,5,5,5,5,5,5,1.5,3.8,1.2,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,3.8,1.8,1.8,1.2,1.2,1.8,3.8,3.8,1.8");
            vbtFunctionBase.SetParamValue("MeasureI_Range", "0.0002");
            instanceRow.ArgList = vbtFunctionBase.Parameters;
            instanceRow.Args = vbtFunctionBase.Args;
            instanceRows.Add(instanceRow);
        }

        public List<SubFlowSheet> GenSubFlowSheets(string idsSheet)
        {
            var subFlowSheets = new List<SubFlowSheet>();
            var flowSheetName = "Flow_" + idsSheet;
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
            foreach (var row in Rows)
                GenFlowRows(row, type, flowRows);
            return flowRows;
        }

        private static void GenFlowRows(PmicIdsSheetRow row, string type, List<FlowRow> flowRows)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = "Test";
            flowRow.Parameter = row.Instance + "_PreSetUp" + "_" + type;
            flowRow.FailAction = "F_" + row.Instance + "_" + type;
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = "Test";
            flowRow.Parameter = row.Instance + "_" + type;
            flowRow.FailAction = "F_" + row.Instance + "_" + type;
            flowRows.Add(flowRow);

            flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = "Bin_" + row.Instance + "_" + type;
            flowRows.Add(flowRow);
        }

        public List<BinTableRow> GenBinTableRows()
        {
            var binTableRows = new List<BinTableRow>();
            foreach (var row in Rows)
                GenBinTableRows(row, binTableRows, "NV");
            foreach (var row in Rows)
                GenBinTableRows(row, binTableRows, "HV");
            foreach (var row in Rows)
                GenBinTableRows(row, binTableRows, "LV");
            return binTableRows;
        }

        private void GenBinTableRows(PmicIdsSheetRow row, List<BinTableRow> binTableRows, string type)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.Pmic, SheetName);
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = "Bin_" + row.Instance + "_" + type;
            binTableRow.ItemList = "F_" + row.Instance + "_" + type;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        #region Field

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<PmicIdsSheetRow> Rows { get; set; }

        public int TestItemIndex = -1;
        public int TtrIndex = -1;
        public int StepIndex = -1;
        public int DescriptionIndex = -1;
        public int InstanceIndex = -1;
        public int MeasurePinsIndex = -1;
        public int CurrentRangeIndex = -1;
        public int Cp1LoLimitIndex = -1;
        public int Cp1HiLimitIndex = -1;
        public int CommentIndex = -1;

        #endregion
    }

    public class PmicIdsReader
    {
        private const string HeaderTestItem = "Test Item";
        private const string HeaderTtr = "TTR";
        private const string HeaderStep = "Step";
        private const string HeaderDescription = "Description";
        private const string HeaderInstance = "Instance";
        private const string HeaderMeasurePins = "MeasurePins";
        private const string HeaderCurrentRange = "CurrentRange";
        private const string HeaderCp1LoLimit = "CP1 Lo Limit (H,L,N)";
        private const string HeaderCp1HiLimit = "CP1 Hi Limit (H,L,N)";
        private const string HeaderComment = "Comment";
        private int _commentIndex = -1;
        private int _cp1HiLimitIndex = -1;
        private int _cp1LoLimitIndex = -1;
        private int _currentRangeIndex = -1;
        private int _descriptionIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _instanceIndex = -1;
        private int _measurePinsIndex = -1;
        private PmicIdsSheetSheet _pmicIdsSheetSheet;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _stepIndex = -1;
        private int _testItemIndex = -1;
        private int _ttrIndex = -1;

        public PmicIdsSheetSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _pmicIdsSheetSheet = new PmicIdsSheetSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _pmicIdsSheetSheet = ReadSheetData();

            return _pmicIdsSheetSheet;
        }

        private PmicIdsSheetSheet ReadSheetData()
        {
            var pmicIdsSheetSheet = new PmicIdsSheetSheet(_sheetName);
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new PmicIdsSheetRow(_sheetName);
                row.RowNum = i;
                if (_testItemIndex != -1)
                    row.TestItem = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _testItemIndex).Trim();
                if (_ttrIndex != -1)
                    row.Ttr = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _ttrIndex).Trim();
                if (_stepIndex != -1)
                    row.Step = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _stepIndex).Trim();
                if (_descriptionIndex != -1)
                    row.Description = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _descriptionIndex).Trim();
                if (_instanceIndex != -1)
                    row.Instance = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _instanceIndex).Trim();
                if (_measurePinsIndex != -1)
                    row.MeasurePins = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _measurePinsIndex).Trim();
                if (_currentRangeIndex != -1)
                    row.CurrentRange = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _currentRangeIndex)
                        .Trim();
                if (_cp1LoLimitIndex != -1)
                    row.Cp1LoLimit = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cp1LoLimitIndex).Trim();
                if (_cp1HiLimitIndex != -1)
                    row.Cp1HiLimit = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _cp1HiLimitIndex).Trim();
                if (_commentIndex != -1)
                    row.Comment = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
                if (!string.IsNullOrEmpty(row.Instance))
                    pmicIdsSheetSheet.Rows.Add(row);
            }

            pmicIdsSheetSheet.TestItemIndex = _testItemIndex;
            pmicIdsSheetSheet.TtrIndex = _ttrIndex;
            pmicIdsSheetSheet.StepIndex = _stepIndex;
            pmicIdsSheetSheet.DescriptionIndex = _descriptionIndex;
            pmicIdsSheetSheet.InstanceIndex = _instanceIndex;
            pmicIdsSheetSheet.MeasurePinsIndex = _measurePinsIndex;
            pmicIdsSheetSheet.CurrentRangeIndex = _currentRangeIndex;
            pmicIdsSheetSheet.Cp1LoLimitIndex = _cp1LoLimitIndex;
            pmicIdsSheetSheet.Cp1HiLimitIndex = _cp1HiLimitIndex;
            pmicIdsSheetSheet.CommentIndex = _commentIndex;

            return pmicIdsSheetSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderTestItem, StringComparison.OrdinalIgnoreCase))
                {
                    _testItemIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTtr, StringComparison.OrdinalIgnoreCase))
                {
                    _ttrIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderStep, StringComparison.OrdinalIgnoreCase))
                {
                    _stepIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDescription, StringComparison.OrdinalIgnoreCase))
                {
                    _descriptionIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderInstance, StringComparison.OrdinalIgnoreCase))
                {
                    _instanceIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderMeasurePins, StringComparison.OrdinalIgnoreCase))
                {
                    _measurePinsIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCurrentRange, StringComparison.OrdinalIgnoreCase))
                {
                    _currentRangeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCp1LoLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _cp1LoLimitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCp1HiLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _cp1HiLimitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderComment, StringComparison.OrdinalIgnoreCase)) _commentIndex = i;
            }


            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
            for (var j = 1; j <= colNum; j++)
                if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim()
                    .Equals(HeaderTestItem, StringComparison.OrdinalIgnoreCase))
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
            _testItemIndex = -1;
            _ttrIndex = -1;
            _stepIndex = -1;
            _descriptionIndex = -1;
            _instanceIndex = -1;
            _measurePinsIndex = -1;
            _currentRangeIndex = -1;
            _cp1LoLimitIndex = -1;
            _cp1HiLimitIndex = -1;
            _commentIndex = -1;
        }
    }
}