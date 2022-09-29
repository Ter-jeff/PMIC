using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.VbtGenTool.Reader
{
    [Serializable]
    public class VbtGenTestPlanRowPlus : VbtGenTestPlanRow
    {
        public VbtGenTestPlanRowPlus(string sourceSheetName)
        {
            VbtGenTestPlanRow = new VbtGenTestPlanRow(sourceSheetName);
        }

        public VbtGenTestPlanRow VbtGenTestPlanRow { get; set; }

        public VbtGenTestPlanRowPlus DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as VbtGenTestPlanRowPlus;
            }
        }
    }

    [Serializable]
    public class VbtGenTestPlanRow
    {
        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string TopList { get; set; }
        public string Command { get; set; }
        public string FunctionName { get; set; }
        public string RegisterMacroName { get; set; }
        public string BitfieldName { get; set; }
        public string Values { get; set; }
        public string Pin { get; set; }
        public string DataLogVariable { get; set; }
        public string Unit { get; set; }
        public string LowLimit { get; set; }
        public string HighLimit { get; set; }
        public string CallbackFunction { get; set; }
        public string VrangeL { get; set; }
        public string IrangeL { get; set; }
        public string VoltageL { get; set; }
        public string CurrentL { get; set; }
        public string VrangeM { get; set; }
        public string IrangeM { get; set; }
        public string VoltageM { get; set; }
        public string CurrentM { get; set; }
        public string Frequency { get; set; }
        public string SampleSize { get; set; }
        public string Comment { get; set; }

        #endregion

        #region Constructor

        public VbtGenTestPlanRow()
        {
            TopList = "";
            Command = "";
            FunctionName = "";
            RegisterMacroName = "";
            BitfieldName = "";
            Values = "";
            Pin = "";
            DataLogVariable = "";
            Unit = "";
            LowLimit = "";
            HighLimit = "";
            CallbackFunction = "";
            VrangeL = "";
            IrangeL = "";
            VoltageL = "";
            CurrentL = "";
            VrangeM = "";
            IrangeM = "";
            VoltageM = "";
            CurrentM = "";
            Frequency = "";
            SampleSize = "";
            Comment = "";
        }

        public VbtGenTestPlanRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            TopList = "";
            Command = "";
            FunctionName = "";
            RegisterMacroName = "";
            BitfieldName = "";
            Values = "";
            Pin = "";
            DataLogVariable = "";
            Unit = "";
            LowLimit = "";
            HighLimit = "";
            CallbackFunction = "";
            VrangeL = "";
            IrangeL = "";
            VoltageL = "";
            CurrentL = "";
            VrangeM = "";
            IrangeM = "";
            VoltageM = "";
            CurrentM = "";
            Frequency = "";
            SampleSize = "";
            Comment = "";
        }

        #endregion
    }

    [Serializable]
    public class VbtGenTestPlanSheet
    {
        #region Constructor

        public VbtGenTestPlanSheet(string sheetName)
        {
            SheetName = sheetName;
            RowList = new List<VbtGenTestPlanRowPlus>();
        }

        #endregion

        public VbtGenTestPlanSheet DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as VbtGenTestPlanSheet;
            }
        }

        #region Property

        public string SheetName { get; set; }
        public int SheetSequence { get; set; }
        public string FullFlowName { get; set; }
        public string FirstFlowName { get; set; }
        public int FlowSequence { get; set; }
        public string ModuleName { get; set; }
        public List<VbtGenTestPlanRowPlus> RowList { get; set; }
        public Dictionary<string, int> HeaderIndex { get; set; } = new Dictionary<string, int>();

        #endregion
    }

    public class VbtGenTestPlanSheetReader
    {
        private const string HeaderTestPlan = "TEST PLAN FOR";
        private const string HeaderTopList = "TOP_LIST";
        private const string HeaderCommand = "COMMAND";
        private const string HeaderFunctionName = "FUNCTION_NAME";
        private const string HeaderRegisterMacroName = "REGISTER/MACRO NAME";
        private const string HeaderBitFieldName = "BITFIELD NAME";
        private const string HeaderValues = "VALUE(S)";
        private const string HeaderPin = "PIN";
        private const string HeaderDataLogVariable = "DATALOG_VARIABLE";
        private const string HeaderUnit = "UNIT";
        private const string HeaderLowLimit = "LOW_LIMIT";
        private const string HeaderHighLimit = "HIGH_LIMIT";
        private const string HeaderCallbackFunction = "CALLBACK FUNCTION";
        private const string HeaderVRangeL = "V_Range_L";
        private const string HeaderIRangeL = "I_Range_L";
        private const string HeaderVoltageL = "Voltage_L";
        private const string HeaderCurrentL = "Current_L";
        private const string HeaderVRangeM = "V_Range_M";
        private const string HeaderIRangeM = "I_Range_M";
        private const string HeaderVoltageM = "Voltage_M";
        private const string HeaderCurrentM = "Current_M";
        private const string HeaderFrequency = "Frequency";
        private const string HeaderSampleSize = "Sample_Size";
        private const string HeaderComment = "COMMENT";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"TOP_LIST", true},
            {"COMMAND", true},
            {"FUNCTION_NAME", true},
            {"REGISTER/MACRO NAME", true},
            {"BITFIELD NAME", true},
            {"VALUE(S)", true},
            {"PIN", true},
            {"DATALOG_VARIABLE", true},
            {"UNIT", true},
            {"LOW_LIMIT", true},
            {"HIGH_LIMIT", true},
            {"CALLBACK FUNCTION", true},
            {"V_Range_L", true},
            {"I_Range_L", true},
            {"Voltage_L", true},
            {"Current_L", true},
            {"V_Range_M", true},
            {"I_Range_M", true},
            {"Voltage_M", true},
            {"Current_M", true},
            {"Frequency", true},
            {"Sample_Size", true},
            {"COMMENT", true}
        };

        private int _bitfieldNameIndex = -1;
        private int _callBackFunctionIndex = -1;
        private int _commandIndex = -1;
        private int _commentIndex = -1;
        private int _currentLIndex = -1;
        private int _currentMIndex = -1;
        private int _datalogVariableIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _frequencyIndex = -1;
        private int _functionNameIndex = -1;
        private int _highLimitIndex = -1;
        private int _iRangeLIndex = -1;
        private int _iRangeMIndex = -1;
        private int _lowLimitIndex = -1;
        private int _pinIndex = -1;
        private int _registerMacroNameIndex = -1;
        private int _sampleSizeIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _topListIndex = -1;
        private int _unitIndex = -1;
        private int _valuesIndex = -1;
        private VbtGenTestPlanSheet _vbtGenTestPlanSheet;
        private int _voltageLIndex = -1;
        private int _voltageMIndex = -1;
        private int _vRangeLIndex = -1;
        private int _vRangeMIndex = -1;

        public VbtGenTestPlanSheet SheetReader(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _vbtGenTestPlanSheet = new VbtGenTestPlanSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetFlowPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _vbtGenTestPlanSheet = ReadSheetData();

            return _vbtGenTestPlanSheet;
        }

        private VbtGenTestPlanSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new VbtGenTestPlanRowPlus(_sheetName);
                row.RowNum = i;
                if (_topListIndex != -1)
                    row.TopList = _excelWorksheet.GetMergedCellValue(i, _topListIndex).Trim();
                if (_commandIndex != -1)
                    row.Command = _excelWorksheet.GetMergedCellValue(i, _commandIndex).Trim();
                if (_functionNameIndex != -1)
                    row.FunctionName = _excelWorksheet.GetMergedCellValue(i, _functionNameIndex)
                        .Trim();
                if (_registerMacroNameIndex != -1)
                    row.RegisterMacroName = _excelWorksheet.GetMergedCellValue(i, _registerMacroNameIndex).Trim();
                if (_bitfieldNameIndex != -1)
                    row.BitfieldName = _excelWorksheet.GetMergedCellValue(i, _bitfieldNameIndex)
                        .Trim();
                if (_valuesIndex != -1)
                    row.Values = _excelWorksheet.GetMergedCellValue(i, _valuesIndex).Trim();
                if (_pinIndex != -1)
                    row.Pin = _excelWorksheet.GetMergedCellValue(i, _pinIndex).Trim();
                if (_datalogVariableIndex != -1)
                    row.DataLogVariable = _excelWorksheet.GetMergedCellValue(i, _datalogVariableIndex)
                        .Trim();
                if (_unitIndex != -1)
                    row.Unit = _excelWorksheet.GetMergedCellValue(i, _unitIndex).Trim();
                if (_lowLimitIndex != -1)
                    row.LowLimit = _excelWorksheet.GetMergedCellValue(i, _lowLimitIndex).Trim();
                if (_highLimitIndex != -1)
                    row.HighLimit = _excelWorksheet.GetMergedCellValue(i, _highLimitIndex).Trim();
                if (_callBackFunctionIndex != -1)
                    row.CallbackFunction = _excelWorksheet.GetMergedCellValue(i, _callBackFunctionIndex).Trim();
                if (_vRangeLIndex != -1)
                    row.VrangeL = _excelWorksheet.GetMergedCellValue(i, _vRangeLIndex).Trim();
                if (_iRangeLIndex != -1)
                    row.IrangeL = _excelWorksheet.GetMergedCellValue(i, _iRangeLIndex).Trim();
                if (_voltageLIndex != -1)
                    row.VoltageL = _excelWorksheet.GetMergedCellValue(i, _voltageLIndex).Trim();
                if (_currentLIndex != -1)
                    row.CurrentL = _excelWorksheet.GetMergedCellValue(i, _currentLIndex).Trim();
                if (_vRangeMIndex != -1)
                    row.VrangeM = _excelWorksheet.GetMergedCellValue(i, _vRangeMIndex).Trim();
                if (_iRangeMIndex != -1)
                    row.IrangeM = _excelWorksheet.GetMergedCellValue(i, _iRangeMIndex).Trim();
                if (_voltageMIndex != -1)
                    row.VoltageM = _excelWorksheet.GetMergedCellValue(i, _voltageMIndex).Trim();
                if (_currentMIndex != -1)
                    row.CurrentM = _excelWorksheet.GetMergedCellValue(i, _currentMIndex).Trim();
                if (_frequencyIndex != -1)
                    row.Frequency = _excelWorksheet.GetMergedCellValue(i, _frequencyIndex).Trim();
                if (_sampleSizeIndex != -1)
                    row.SampleSize = _excelWorksheet.GetMergedCellValue(i, _sampleSizeIndex).Trim();
                if (_commentIndex != -1)
                    row.Comment = _excelWorksheet.GetMergedCellValue(i, _commentIndex).Trim();
                _vbtGenTestPlanSheet.RowList.Add(row);
            }

            return _vbtGenTestPlanSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderTopList, StringComparison.OrdinalIgnoreCase))
                {
                    _topListIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderTopList, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCommand, StringComparison.OrdinalIgnoreCase))
                {
                    _commandIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderCommand, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFunctionName, StringComparison.OrdinalIgnoreCase))
                {
                    _functionNameIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderFunctionName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderRegisterMacroName, StringComparison.OrdinalIgnoreCase))
                {
                    _registerMacroNameIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderRegisterMacroName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderBitFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    _bitfieldNameIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderBitFieldName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderValues, StringComparison.OrdinalIgnoreCase))
                {
                    _valuesIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderValues, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderPin, StringComparison.OrdinalIgnoreCase))
                {
                    _pinIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderPin, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderDataLogVariable, StringComparison.OrdinalIgnoreCase))
                {
                    _datalogVariableIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderDataLogVariable, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderUnit, StringComparison.OrdinalIgnoreCase))
                {
                    _unitIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderUnit, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderLowLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _lowLimitIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderLowLimit, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderHighLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _highLimitIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderHighLimit, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCallbackFunction, StringComparison.OrdinalIgnoreCase))
                {
                    _callBackFunctionIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderCallbackFunction, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderVRangeL, StringComparison.OrdinalIgnoreCase))
                {
                    _vRangeLIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderVRangeL, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderIRangeL, StringComparison.OrdinalIgnoreCase))
                {
                    _iRangeLIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderIRangeL, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderVoltageL, StringComparison.OrdinalIgnoreCase))
                {
                    _voltageLIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderVoltageL, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCurrentL, StringComparison.OrdinalIgnoreCase))
                {
                    _currentLIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderCurrentL, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderVRangeM, StringComparison.OrdinalIgnoreCase))
                {
                    _vRangeMIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderVRangeM, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderIRangeM, StringComparison.OrdinalIgnoreCase))
                {
                    _iRangeMIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderIRangeM, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderVoltageM, StringComparison.OrdinalIgnoreCase))
                {
                    _voltageMIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderVoltageM, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderCurrentM, StringComparison.OrdinalIgnoreCase))
                {
                    _currentMIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderCurrentM, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFrequency, StringComparison.OrdinalIgnoreCase))
                {
                    _frequencyIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderFrequency, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderSampleSize, StringComparison.OrdinalIgnoreCase))
                {
                    _sampleSizeIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderSampleSize, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    _vbtGenTestPlanSheet.HeaderIndex.Add(HeaderComment, i);
                }
            }

            foreach (var header in _vbtGenTestPlanSheet.HeaderIndex)
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                    return false;

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i < rowNum; i++)
                for (var j = 1; j < colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                            .Equals(HeaderTopList, StringComparison.OrdinalIgnoreCase) &&
                            _excelWorksheet.GetCellValue(i, j + 1).Trim()
                            .Equals(HeaderCommand, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }

            return false;
        }

        private bool GetFlowPosition()
        {
            const string reg1 = @"\<(?<module>.+)\>(?<Flow>.+)";
            const string reg2 = @"_(?<Sequence>\d+)$";

            if (Regex.IsMatch(_vbtGenTestPlanSheet.SheetName, reg2, RegexOptions.IgnoreCase))
                _vbtGenTestPlanSheet.SheetSequence = int.Parse(Regex
                    .Match(_vbtGenTestPlanSheet.SheetName, reg2, RegexOptions.IgnoreCase).Groups["Sequence"]
                    .ToString());

            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i < rowNum; i++)
                for (var j = 1; j < colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                        .StartsWith(HeaderTestPlan, StringComparison.OrdinalIgnoreCase))
                    {
                        var flow = Regex
                            .Match(_excelWorksheet.GetCellValue(i, j).Trim(), reg1,
                                RegexOptions.IgnoreCase).Groups["Flow"].ToString();
                        var module = Regex
                            .Match(_excelWorksheet.GetCellValue(i, j).Trim(), reg1,
                                RegexOptions.IgnoreCase).Groups["module"].ToString();
                        if (!string.IsNullOrEmpty(flow) && !string.IsNullOrEmpty(module))
                        {
                            _vbtGenTestPlanSheet.FullFlowName =
                                flow.ToUpper().LastIndexOf("_FLOW", StringComparison.Ordinal) == -1
                                    ? flow.ToUpper()
                                    : flow.Substring(0, flow.ToUpper().LastIndexOf("_FLOW", StringComparison.Ordinal))
                                        .ToUpper();
                            _vbtGenTestPlanSheet.FirstFlowName = _vbtGenTestPlanSheet.FullFlowName.Split('_').First();
                            _vbtGenTestPlanSheet.FlowSequence = Regex.IsMatch(flow, reg2, RegexOptions.IgnoreCase)
                                ? int.Parse(Regex.Match(flow, reg2, RegexOptions.IgnoreCase).Groups["Sequence"].ToString())
                                : 0;
                            _vbtGenTestPlanSheet.ModuleName = module.ToUpper();
                            return true;
                        }

                        return false;
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
            _topListIndex = -1;
            _commandIndex = -1;
            _functionNameIndex = -1;
            _registerMacroNameIndex = -1;
            _bitfieldNameIndex = -1;
            _valuesIndex = -1;
            _pinIndex = -1;
            _datalogVariableIndex = -1;
            _unitIndex = -1;
            _lowLimitIndex = -1;
            _highLimitIndex = -1;
            _callBackFunctionIndex = -1;
            _vRangeLIndex = -1;
            _iRangeLIndex = -1;
            _voltageLIndex = -1;
            _currentLIndex = -1;
            _vRangeMIndex = -1;
            _iRangeMIndex = -1;
            _voltageMIndex = -1;
            _currentMIndex = -1;
            _frequencyIndex = -1;
            _sampleSizeIndex = -1;
            _commentIndex = -1;
        }
    }
}