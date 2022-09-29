//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS}
//
//------------------------------------------------------------------------------ 

using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.VbtGenTool.Reader
{
    public class TestParameterRow
    {
        #region Constructor

        public TestParameterRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
        }

        #endregion

        #region Property

        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string FunctionName { set; get; }
        public string BlockName { set; get; }
        public string TrimMethod { set; get; }
        public string TrimTable { set; get; }
        public string CodeDistribution { set; get; }
        public string NumBits { set; get; }
        public string Lsb { set; get; }
        public string Target { set; get; }
        public string PreTrimCode { set; get; }
        public string TrimRegister { set; get; }
        public string TrimBitField { set; get; }
        public string OtpRegister { set; get; }
        public string MeasPin { set; get; }
        public string WaitTime { set; get; }
        public string SampleSize { set; get; }
        public string DatalogTemplate { set; get; }
        public string PowerPin { set; get; }
        public string ToggleDirection { set; get; }
        public string ToggleThresholdFailCount { set; get; }
        public string AnalogSweepPin { set; get; }
        public string AnalogSweepStart { set; get; }
        public string AnalogSweepStop { set; get; }
        public string AnalogSweepStep { set; get; }
        public string TrimLinkSweepStart { set; get; }
        public string TrimLinkSweepStop { set; get; }
        public string GngLow { set; get; }
        public string GngHigh { set; get; }
        public string LowLimit { set; get; }
        public string HighLimit { set; get; }

        #endregion
    }

    public class TestParameterSheet
    {
        private const string FwTrim = "FWTrim";
        private const string FwDtbAnaI = "FW_DTB_AnaI";
        private const string FwDtbAnaV = "FW_DTB_AnaV";
        private const string FwAtb3StepTrim = "FW_ATB_3StepTrim";

        public List<string> VoltagesType = new List<string> { "", "_LV", "_HV" };

        #region Constructor

        public TestParameterSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<TestParameterRow>();
        }

        #endregion

        public InstanceSheet GenInstance()
        {
            var instanceSheet = new InstanceSheet("TestInst_" + Block);
            instanceSheet.AddHeaderFooter(Block + "_Trim");
            foreach (var voltage in VoltagesType)
                instanceSheet.AddHeaderFooter(Block + "_PostBurn" + voltage);

            var trimInsRowlst = new List<InstanceRow>();
            var postBurnInsRowlst = new List<InstanceRow>();
            var measureInsRowlst = new List<InstanceRow>();
            foreach (var row in Rows)
            {
                var functionName = row.FunctionName;
                var instance = GenInstanceRow(functionName, "");
                if (row.TrimMethod.Equals(FwTrim, StringComparison.CurrentCultureIgnoreCase))
                {
                    instance.ArgList = "bIsTRIM";
                    instance.Args[0] = "TRUE";
                    trimInsRowlst.Add(instance);

                    var testName = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                        ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                        : functionName + "_PostBurn";
                    var postBurn = instance.DeepClone();
                    postBurn.TestName = testName;
                    postBurn.Args[0] = "FALSE";
                    postBurnInsRowlst.Add(postBurn);
                }
                else if (row.TrimMethod.Equals(FwDtbAnaI, StringComparison.CurrentCultureIgnoreCase) ||
                         row.TrimMethod.Equals(FwDtbAnaV, StringComparison.CurrentCultureIgnoreCase))
                {
                    instance.ArgList = "bIsTRIM";
                    instance.Args[0] = "TRUE";
                    trimInsRowlst.Add(instance);

                    var testName = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                        ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                        : functionName + "_PostBurn";
                    var postBurn = instance.DeepClone();
                    postBurn.TestName = testName;
                    postBurn.Args[0] = "FALSE";
                    postBurnInsRowlst.Add(postBurn);

                    var meas = instance.DeepClone();
                    meas.Name = functionName + "_Meas";
                    meas.Args.Clear();
                    meas.TestName = functionName + "_Meas";
                    measureInsRowlst.Add(meas);

                    var testName2 = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                        ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                        : functionName + "_PostBurn";
                    var measPost = instance.DeepClone();
                    measPost.Name = functionName + "_Meas";
                    measPost.Args.Clear();
                    measPost.TestName = testName2 + "_Meas";
                    measureInsRowlst.Add(measPost);
                }
                else if (row.TrimMethod.Equals(FwAtb3StepTrim, StringComparison.CurrentCultureIgnoreCase) ||
                         row.TrimMethod.Equals("3StepTrim", StringComparison.CurrentCultureIgnoreCase) ||
                         row.TrimMethod.Equals("BestCodeSearch", StringComparison.CurrentCultureIgnoreCase) ||
                         row.TrimMethod.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
                {
                    instance.ArgList = "bIsTRIM, bPreTrimCheck, bEnableTrimlink";
                    instance.Args[0] = "TRUE";
                    instance.Args[1] = "FALSE";
                    instance.Args[2] = "FALSE";
                    trimInsRowlst.Add(instance);

                    var testName = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                        ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                        : functionName + "_PostBurn";
                    var postBurn = instance.DeepClone();
                    postBurn.TestName = testName;
                    postBurn.Args[0] = "FALSE";
                    postBurn.Args[1] = "FALSE";
                    instance.Args[2] = "FALSE";
                    postBurnInsRowlst.Add(postBurn);
                }
                else //if (row.TrimMethod.Equals("CodeSweep", StringComparison.CurrentCultureIgnoreCase))
                {
                    // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Add Start
                    if (row.TrimMethod.Equals("FWTrim", StringComparison.CurrentCultureIgnoreCase))
                    {
                        instance.ArgList = "bIsTRIM";
                        instance.Args[0] = "TRUE";
                        trimInsRowlst.Add(instance);
                        var testName = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                            ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                            : functionName + "_PostBurn";
                        var postBurn = instance.DeepClone();
                        postBurn.TestName = testName;
                        postBurn.Args[0] = "FALSE";
                        postBurnInsRowlst.Add(postBurn);
                    }
                    else
                    {
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Add End
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Chg Start
                        //instance.ArgList = "bPreTrimCheck, bIsTRIM, bIsSweepCode, bEnSweepAnalog, bFWTrim"; // "bIsTRIM, bIsSweepCode, bFWTrim, Validating_";
                        instance.ArgList =
                            "bPreTrimCheck, bIsTRIM, bIsSweepCode, bEnSweepAnalog"; // "bIsTRIM, bIsSweepCode, bFWTrim, Validating_";
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Chg End
                        //if (Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase))
                        //{
                        instance.Args[0] = "FALSE";
                        instance.Args[1] = "TRUE";
                        instance.Args[2] = "TRUE";
                        instance.Args[3] = "TRUE";
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Del Start
                        //instance.Args[4] = "FALSE";
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Del End
                        trimInsRowlst.Add(instance);
                        //}

                        //var testName = Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase);
                        var testName = Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase)
                            ? Regex.Replace(functionName, "_Trim", "_PostBurn", RegexOptions.IgnoreCase)
                            : functionName + "_PostBurn";
                        var postBurn = instance.DeepClone();
                        postBurn.TestName = testName;
                        postBurn.Args[0] = "FALSE";
                        postBurn.Args[1] = "FALSE";
                        postBurn.Args[2] = "FALSE";
                        postBurn.Args[3] = "TRUE";
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Del Start
                        //postBurn.Args[4] = "FALSE";
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Del End
                        postBurnInsRowlst.Add(postBurn);
                        // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Add Start
                    }
                    // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Add End
                }
            }

            instanceSheet.AddRows(trimInsRowlst);
            instanceSheet.AddRows(postBurnInsRowlst);
            instanceSheet.AddRows(measureInsRowlst);
            instanceSheet.InstanceRows = instanceSheet.InstanceRows.Distinct(new InstanceRowComparer()).ToList();
            return instanceSheet;
        }

        public virtual InstanceRow GenInstanceRow(string functionName, string timeSetName)
        {
            var vbtFunction = TestProgram.VbtFunctionLib.GetFunctionByName(functionName);
            var instanceRow = new InstanceRow();
            instanceRow.TestName = functionName;
            instanceRow.Type = "VBT";
            instanceRow.Name = functionName;
            instanceRow.DcCategory = "Analog";
            instanceRow.DcSelector = "Typ";
            instanceRow.TimeSets = timeSetName;
            instanceRow.AcCategory = "Common"; //CreateAcCategory(timeSetName);
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = "Levels_Analog";
            vbtFunction.SetParamValue("AHB_WRITE_OPTION", "-1");
            vbtFunction.SetParamValue("FLAT_PATTERN_OPTION", "0");
            instanceRow.ArgList = vbtFunction.Parameters;
            instanceRow.Args = vbtFunction.Args;
            return instanceRow;
        }

        public List<SubFlowSheet> GenFlowSheet()
        {
            var subFlowSheets = new List<SubFlowSheet>();

            var flowNameTrim = Block + "_Trim";

            var subFlowTrim = new SubFlowSheet("Flow_" + flowNameTrim);

            subFlowTrim.AddRow(GenSetErrorBin());

            subFlowTrim.FlowRows.AddStartRows(subFlowTrim.SheetName, SubFlowSheet.Ttime);

            foreach (var row in Rows)
            {
                subFlowTrim.AddRow(GenFlow(row, "_Trim", flowNameTrim));
                subFlowTrim.FlowRows.Add_A_Enable_MP_SBIN(Block);
                subFlowTrim.AddRow(GenBinTable(row, "_Trim", flowNameTrim));

                if (row.TrimMethod.Equals(FwDtbAnaI, StringComparison.CurrentCultureIgnoreCase) ||
                    row.TrimMethod.Equals(FwDtbAnaV, StringComparison.CurrentCultureIgnoreCase))
                {
                    subFlowTrim.AddRow(GenFlowMeas(row, "_Trim", flowNameTrim));
                    subFlowTrim.FlowRows.Add_A_Enable_MP_SBIN(Block);
                    subFlowTrim.AddRow(GenBinTableMeas(row, "_Trim", flowNameTrim));
                }
            }

            subFlowTrim.FlowRows.AddEndRows(subFlowTrim.SheetName, SubFlowSheet.Ttime);
            //subFlowTrim.FlowRows = subFlowTrim.FlowRows.Distinct(new FlowRowComparer()).ToList();
            subFlowSheets.Add(subFlowTrim);

            foreach (var voltage in VoltagesType)
            {
                var flowName = Block + "_PostBurn" + voltage;

                var subFlowPostBurn = new SubFlowSheet("Flow_" + Block + "_PostBurn" + voltage);

                subFlowPostBurn.AddRow(GenSetErrorBin());

                subFlowPostBurn.FlowRows.AddStartRows(subFlowPostBurn.SheetName, SubFlowSheet.Ttime);

                foreach (var row in Rows)
                {
                    if (string.IsNullOrEmpty(voltage))
                        subFlowPostBurn.AddRow(GenFlow(row, "_PostBurn", flowName));
                    else
                        subFlowPostBurn.AddRow(GenFlow(row, "_PostBurn", flowName, false));

                    if (string.IsNullOrEmpty(voltage))
                    {
                        subFlowPostBurn.FlowRows.Add_A_Enable_MP_SBIN(Block);
                        subFlowPostBurn.AddRow(GenBinTable(row, "_PostBurn", flowName));
                    }

                    if (row.TrimMethod.Equals(FwDtbAnaI, StringComparison.CurrentCultureIgnoreCase) ||
                        row.TrimMethod.Equals(FwDtbAnaV, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(voltage))
                            subFlowPostBurn.AddRow(GenFlowMeas(row, "_PostBurn", flowName));
                        else
                            subFlowPostBurn.AddRow(GenFlowMeas(row, "_PostBurn", flowName, false));

                        if (string.IsNullOrEmpty(voltage))
                        {
                            subFlowPostBurn.FlowRows.Add_A_Enable_MP_SBIN(Block);
                            subFlowPostBurn.AddRow(GenBinTableMeas(row, "_PostBurn", flowName));
                        }
                    }
                }

                subFlowPostBurn.FlowRows.AddEndRows(subFlowPostBurn.SheetName, SubFlowSheet.Ttime);
                //subFlow.FlowRows = subFlow.FlowRows.Distinct(new FlowRowComparer()).ToList();
                subFlowSheets.Add(subFlowPostBurn);
            }

            return subFlowSheets;
        }

        private FlowRow GenFlow(TestParameterRow row, string name, string flowName, bool isFailFlag = true)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeTest;
            flowRow.Parameter = GetParameter(row.FunctionName, name);
            if (isFailFlag)
                flowRow.FailAction = GetFailFlag(row.FunctionName, name, flowName).AddBlockFlag(Block);
            return flowRow;
        }

        private FlowRow GenFlowMeas(TestParameterRow row, string name, string flowName, bool isFailFlag = true)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeTest;
            flowRow.Parameter = GetParameter(row.FunctionName, name) + "_Meas";
            if (isFailFlag)
                flowRow.FailAction = (GetFailFlag(row.FunctionName, name, flowName) + "_Meas").AddBlockFlag(Block);
            return flowRow;
        }

        private FlowRow GenBinTableMeas(TestParameterRow row, string name, string sheetName)
        {
            var rowBin = new FlowRow();
            rowBin.OpCode = FlowRow.OpCodeBinTable;
            rowBin.Parameter = GetBinTableName(row.FunctionName, name, sheetName) + "_Meas";
            var last = sheetName.Split('_').Last();
            if (!string.IsNullOrEmpty(last))
                if (VoltagesType.Exists(x => x.Equals("_" + last, StringComparison.CurrentCultureIgnoreCase)))
                    rowBin.Env = "X";
            return rowBin;
        }

        private FlowRow GenBinTable(TestParameterRow row, string name, string sheetName)
        {
            var rowBin = new FlowRow();
            rowBin.OpCode = FlowRow.OpCodeBinTable;
            rowBin.Parameter = GetBinTableName(row.FunctionName, name, sheetName);
            var last = sheetName.Split('_').Last();
            if (!string.IsNullOrEmpty(last))
                if (VoltagesType.Exists(x => x.Equals("_" + last, StringComparison.CurrentCultureIgnoreCase)))
                    rowBin.Env = "X";
            return rowBin;
        }

        private string GetParameter(string functionName, string replace)
        {
            if (Regex.IsMatch(functionName, "_Trim", RegexOptions.IgnoreCase))
                functionName = Regex.Replace(functionName, "_Trim", replace, RegexOptions.IgnoreCase);
            else if (replace.Equals("_PostBurn"))
                functionName = functionName + "_PostBurn";
            return functionName;
        }

        private string GetBinTableName(string functionName, string replace, string sheetName)
        {
            return "Bin_" + GetParameter(functionName, replace);
        }

        private string GetFailFlag(string functionName, string replace, string sheetName)
        {
            // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Chg Start
            //return "F_" + GetParameter(functionName, replace) + "_" + sheetName;
            if (replace == "_Trim")
            {
                // _Trim
                if (!functionName.EndsWith("_Trim", StringComparison.OrdinalIgnoreCase)) functionName += "_Trim";
            }
            else
            {
                // _PostBurn
                if (functionName.EndsWith("_Trim", StringComparison.OrdinalIgnoreCase))
                    functionName = functionName.Substring(0, functionName.Length - 5);
                if (!functionName.EndsWith("_PostBurn", StringComparison.OrdinalIgnoreCase))
                    functionName += "_PostBurn";
            }

            return "F_" + functionName;
            // 2021-06-21  Bruce Qian     #77             Remove Start_Of_Test argument bFWTrim in {PARAM_DEFINITIONS} Chg End
        }

        protected FlowRow GenSetErrorBin()
        {
            var rowBin = new FlowRow();
            rowBin.OpCode = "set-error-bin";
            rowBin.BinFail = "999";
            rowBin.SortFail = "999";
            return rowBin;
        }

        public BinTableRows GenTrimBinTableRows()
        {
            var trimBinTableRowlst = new BinTableRows();
            foreach (var row in Rows)
            //Trim
            {
                var sheetName = Block + "_Trim";
                var para = new BinNumberRuleCondition(EnumBinNumberBlock.Pmic, sheetName);
                BinNumberRuleRow bin;
                BinNumberSingleton.Instance().GetBinNumDefRow(para, out bin);
                var binRow = new BinTableRow();
                var name = row.FunctionName;
                binRow.Name = GetBinTableName(name, "_Trim", sheetName);
                binRow.ItemList = GetFailFlag(row.FunctionName, "_Trim", sheetName);
                binRow.Op = "AND";
                binRow.Items = Enumerable.Repeat("T", 1).ToList();
                binRow.Result = "Fail";
                binRow.Sort = "9999"; //bin.CurrentSoftBin.ToString();
                binRow.Bin = "9"; //bin.HardBin;
                if (!trimBinTableRowlst.Exists(x => x.Name.Equals(binRow.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    trimBinTableRowlst.Add(binRow);

                    if (row.TrimMethod.Equals(FwDtbAnaI, StringComparison.CurrentCultureIgnoreCase) ||
                        row.TrimMethod.Equals(FwDtbAnaV, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var meas = binRow.DeepClone();
                        meas.Name = GetBinTableName(name, "_Trim", sheetName) + "_Meas";
                        meas.ItemList = GetFailFlag(row.FunctionName, "_Trim", sheetName) + "_Meas";
                        trimBinTableRowlst.Add(meas);
                    }
                }
            }

            var distinct = trimBinTableRowlst.Distinct(new BinTableRowComparer()).ToList();
            trimBinTableRowlst = new BinTableRows(distinct);
            return trimBinTableRowlst;
        }

        public BinTableRows GenPostBurnBinTableRows()
        {
            var postBurnBinTableRowlst = new BinTableRows();
            foreach (var row in Rows)
            //PostBurn
            //foreach (var voltage in Voltages)
            {
                var binRow = new BinTableRow();
                var sheetName = Block + "_PostBurn";
                //if (!string.IsNullOrEmpty(voltage))
                //    sheetName = Block + "_PostBurn" + voltage;
                var para = new BinNumberRuleCondition(EnumBinNumberBlock.Pmic, sheetName);
                BinNumberRuleRow bin;
                BinNumberSingleton.Instance().GetBinNumDefRow(para, out bin);
                var name = row.FunctionName;
                binRow.Name = GetBinTableName(name, "_PostBurn", sheetName);
                binRow.ItemList = GetFailFlag(row.FunctionName, "_PostBurn", sheetName);
                binRow.Op = "AND";
                binRow.Items = Enumerable.Repeat("T", 1).ToList();
                binRow.Result = "Fail";
                binRow.Sort = "9999"; //bin.CurrentSoftBin.ToString();
                binRow.Bin = "9"; //bin.HardBin;
                if (!postBurnBinTableRowlst.Exists(x => x.Name.Equals(binRow.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    postBurnBinTableRowlst.Add(binRow);
                    if (row.TrimMethod.Equals(FwDtbAnaI, StringComparison.CurrentCultureIgnoreCase) ||
                        row.TrimMethod.Equals(FwDtbAnaV, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var meas = binRow.DeepClone();
                        meas.Name = GetBinTableName(name, "_PostBurn", sheetName) + "_Meas";
                        meas.ItemList = GetFailFlag(row.FunctionName, "_PostBurn", sheetName) + "_Meas";
                        postBurnBinTableRowlst.Add(meas);
                    }
                }
            }

            var distinct = postBurnBinTableRowlst.Distinct(new BinTableRowComparer()).ToList();
            postBurnBinTableRowlst = new BinTableRows(distinct);
            return postBurnBinTableRowlst;
        }

        public BinTableRows GenBinTableRows(string block)
        {
            var binTableRows = new BinTableRows();
            binTableRows.GenBlockBinTable(block);
            var postBurnBinTableRows = GenPostBurnBinTableRows();
            binTableRows.AddRange(GenTrimBinTableRows());
            binTableRows.GenSetError(block + "_Trim");
            binTableRows.AddRange(postBurnBinTableRows);
            binTableRows.GenSetError(block + "_PostBurn");
            return binTableRows;
        }

        #region Property

        public string Block { get; set; }
        public string SheetName { get; set; }
        public List<TestParameterRow> Rows { get; set; }

        #endregion
    }

    public class TestParameterReader
    {
        private const string ConHeaderFunctionName = "FunctionName";
        private const string ConHeaderBlockName = "BlockName";
        private const string ConHeaderTrimMethod = "TrimMethod";
        private const string ConHeaderTrimTable = "TrimTable";
        private const string ConHeaderCodeDistribution = "CodeDistribution";
        private const string ConHeaderNumBits = "Numbits";
        private const string ConHeaderLsb = "LSB";
        private const string ConHeaderTarget = "Target";
        private const string ConHeaderPreTrimCode = "PreTrimCode";
        private const string ConHeaderTrimRegister = "TrimRegister";
        private const string ConHeaderTrimBitField = "TrimBitField";
        private const string ConHeaderOtpRegister = "OTPRegister";
        private const string ConHeaderMeasPin = "MeasPin";
        private const string ConHeaderWaitTime = "WaitTime";
        private const string ConHeaderSampleSize = "SampleSize";
        private const string ConHeaderDatalogTemplate = "DatalogTemplate";
        private const string ConHeaderPowerPin = "PowerPin";
        private const string ConHeaderToggleDirection = "ToggleDirection";
        private const string ConHeaderToggleThresholdFailCount = "ToggleThreshold / FailCount";
        private const string ConHeaderAnalogSweepPin = "AnalogSweepPin";
        private const string ConHeaderAnalogSweepStart = "AnalogSweepStart";
        private const string ConHeaderAnalogSweepStop = "AnalogSweepStop";
        private const string ConHeaderAnalogSweepStep = "AnalogSweepStep";
        private const string ConHeaderTrimLinkSweepStart = "TrimLinkSweepStart";
        private const string ConHeaderTrimLinkSweepStop = "TrimLinkSweepStop";
        private const string ConHeaderGngLow = "GNGLow";
        private const string ConHeaderGngHigh = "GNGHigh";
        private const string ConHeaderLowLimit = "LowLimit";
        private const string ConHeaderHighLimit = "HighLimit";
        private int _analogSweepPinIndex = -1;
        private int _analogSweepStartIndex = -1;
        private int _analogSweepStepIndex = -1;
        private int _analogSweepStopIndex = -1;
        private int _blockNameIndex = -1;
        private int _codeDistributionIndex = -1;
        private int _datalogTemplateIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _functionNameIndex = -1;
        private int _gNgHighIndex = -1;
        private int _gNgLowIndex = -1;
        private int _highLimitIndex = -1;
        private int _lowLimitIndex = -1;
        private int _lSbIndex = -1;
        private int _measPinIndex = -1;
        private int _numBitsIndex = -1;
        private int _oTpRegisterIndex = -1;
        private int _powerPinIndex = -1;
        private int _preTrimCodeIndex = -1;
        private int _sampleSizeIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _targetIndex = -1;
        private TestParameterSheet _testParameterSheet;
        private int _toggleDirectionIndex = -1;
        private int _toggleThresholdFailCountIndex = -1;
        private int _trimBitFieldIndex = -1;
        private int _trimLinkSweepStartIndex = -1;
        private int _trimLinkSweepStopIndex = -1;
        private int _trimMethodIndex = -1;
        private int _trimRegisterIndex = -1;
        private int _trimTableIndex = -1;
        private int _waitTimeIndex = -1;

        public TestParameterSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _testParameterSheet = new TestParameterSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _testParameterSheet = ReadSheetData();

            if (StaticTestPlan.VddLevelsSheet != null && StaticTestPlan.VddLevelsSheet.Rows.Any())
                if (StaticTestPlan.VddLevelsSheet.Rows.First().ExtraSelectors.Any())
                    foreach (var extraSelector in StaticTestPlan.VddLevelsSheet.Rows.First().ExtraSelectors)
                        _testParameterSheet.VoltagesType.Add("_" + extraSelector.Key);

            return _testParameterSheet;
        }

        private TestParameterSheet ReadSheetData()
        {
            var a1 = _excelWorksheet.GetMergedCellValue(1, 1);
            if (a1.Contains(">"))
                _testParameterSheet.Block = a1.Substring(a1.IndexOf(">", StringComparison.Ordinal) + 1).Trim();
            if (string.IsNullOrEmpty(_testParameterSheet.Block))
            {
                var error = string.Format("The format of A1 cell - {0} is not correct !!!", a1);
                Response.Report(error, EnumMessageLevel.Error, 0);
                ErrorManager.AddError(EnumErrorType.MissingPattern, EnumErrorLevel.Error, _excelWorksheet.Name, 1, 1,
                    error);
            }

            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new TestParameterRow(_sheetName);
                row.RowNum = i;
                if (_functionNameIndex != -1)
                    row.FunctionName = _excelWorksheet.GetMergedCellValue(i, _functionNameIndex)
                        .Trim();
                if (_blockNameIndex != -1)
                    row.BlockName = _excelWorksheet.GetMergedCellValue(i, _blockNameIndex).Trim();
                if (_trimMethodIndex != -1)
                    row.TrimMethod = _excelWorksheet.GetMergedCellValue(i, _trimMethodIndex).Trim();
                if (_trimTableIndex != -1)
                    row.TrimTable = _excelWorksheet.GetMergedCellValue(i, _trimTableIndex).Trim();
                if (_codeDistributionIndex != -1)
                    row.CodeDistribution = _excelWorksheet.GetMergedCellValue(i, _codeDistributionIndex).Trim();
                if (_numBitsIndex != -1)
                    row.NumBits = _excelWorksheet.GetMergedCellValue(i, _numBitsIndex).Trim();
                if (_lSbIndex != -1)
                    row.Lsb = _excelWorksheet.GetMergedCellValue(i, _lSbIndex).Trim();
                if (_targetIndex != -1)
                    row.Target = _excelWorksheet.GetMergedCellValue(i, _targetIndex).Trim();
                if (_preTrimCodeIndex != -1)
                    row.PreTrimCode = _excelWorksheet.GetMergedCellValue(i, _preTrimCodeIndex).Trim();
                if (_trimRegisterIndex != -1)
                    row.TrimRegister = _excelWorksheet.GetMergedCellValue(i, _trimRegisterIndex)
                        .Trim();
                if (_trimBitFieldIndex != -1)
                    row.TrimBitField = _excelWorksheet.GetMergedCellValue(i, _trimBitFieldIndex)
                        .Trim();
                if (_oTpRegisterIndex != -1)
                    row.OtpRegister = _excelWorksheet.GetMergedCellValue(i, _oTpRegisterIndex).Trim();
                if (_measPinIndex != -1)
                    row.MeasPin = _excelWorksheet.GetMergedCellValue(i, _measPinIndex).Trim();
                if (_waitTimeIndex != -1)
                    row.WaitTime = _excelWorksheet.GetMergedCellValue(i, _waitTimeIndex).Trim();
                if (_sampleSizeIndex != -1)
                    row.SampleSize = _excelWorksheet.GetMergedCellValue(i, _sampleSizeIndex).Trim();
                if (_datalogTemplateIndex != -1)
                    row.DatalogTemplate = _excelWorksheet.GetMergedCellValue(i, _datalogTemplateIndex)
                        .Trim();
                if (_powerPinIndex != -1)
                    row.PowerPin = _excelWorksheet.GetMergedCellValue(i, _powerPinIndex).Trim();
                if (_toggleDirectionIndex != -1)
                    row.ToggleDirection = _excelWorksheet.GetMergedCellValue(i, _toggleDirectionIndex)
                        .Trim();
                if (_toggleThresholdFailCountIndex != -1)
                    row.ToggleThresholdFailCount = _excelWorksheet.GetMergedCellValue(i, _toggleThresholdFailCountIndex).Trim();
                if (_analogSweepPinIndex != -1)
                    row.AnalogSweepPin = _excelWorksheet.GetMergedCellValue(i, _analogSweepPinIndex)
                        .Trim();
                if (_analogSweepStartIndex != -1)
                    row.AnalogSweepStart = _excelWorksheet.GetMergedCellValue(i, _analogSweepStartIndex).Trim();
                if (_analogSweepStopIndex != -1)
                    row.AnalogSweepStop = _excelWorksheet.GetMergedCellValue(i, _analogSweepStopIndex)
                        .Trim();
                if (_analogSweepStepIndex != -1)
                    row.AnalogSweepStep = _excelWorksheet.GetMergedCellValue(i, _analogSweepStepIndex)
                        .Trim();
                if (_trimLinkSweepStartIndex != -1)
                    row.TrimLinkSweepStart = _excelWorksheet.GetMergedCellValue(i, _trimLinkSweepStartIndex).Trim();
                if (_trimLinkSweepStopIndex != -1)
                    row.TrimLinkSweepStop = _excelWorksheet.GetMergedCellValue(i, _trimLinkSweepStopIndex).Trim();
                if (_gNgLowIndex != -1)
                    row.GngLow = _excelWorksheet.GetMergedCellValue(i, _gNgLowIndex).Trim();
                if (_gNgHighIndex != -1)
                    row.GngHigh = _excelWorksheet.GetMergedCellValue(i, _gNgHighIndex).Trim();
                if (_lowLimitIndex != -1)
                    row.LowLimit = _excelWorksheet.GetMergedCellValue(i, _lowLimitIndex).Trim();
                if (_highLimitIndex != -1)
                    row.HighLimit = _excelWorksheet.GetMergedCellValue(i, _highLimitIndex).Trim();
                if (string.IsNullOrEmpty(row.FunctionName))
                    break;
                _testParameterSheet.Rows.Add(row);
            }

            return _testParameterSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderFunctionName, StringComparison.OrdinalIgnoreCase))
                {
                    _functionNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderBlockName, StringComparison.OrdinalIgnoreCase))
                {
                    _blockNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimMethod, StringComparison.OrdinalIgnoreCase))
                {
                    _trimMethodIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimTable, StringComparison.OrdinalIgnoreCase))
                {
                    _trimTableIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCodeDistribution, StringComparison.OrdinalIgnoreCase))
                {
                    _codeDistributionIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderNumBits, StringComparison.OrdinalIgnoreCase))
                {
                    _numBitsIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderLsb, StringComparison.OrdinalIgnoreCase))
                {
                    _lSbIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTarget, StringComparison.OrdinalIgnoreCase))
                {
                    _targetIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPreTrimCode, StringComparison.OrdinalIgnoreCase))
                {
                    _preTrimCodeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimRegister, StringComparison.OrdinalIgnoreCase))
                {
                    _trimRegisterIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimBitField, StringComparison.OrdinalIgnoreCase))
                {
                    _trimBitFieldIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderOtpRegister, StringComparison.OrdinalIgnoreCase))
                {
                    _oTpRegisterIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderMeasPin, StringComparison.OrdinalIgnoreCase))
                {
                    _measPinIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderWaitTime, StringComparison.OrdinalIgnoreCase))
                {
                    _waitTimeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderSampleSize, StringComparison.OrdinalIgnoreCase))
                {
                    _sampleSizeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderDatalogTemplate, StringComparison.OrdinalIgnoreCase))
                {
                    _datalogTemplateIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPowerPin, StringComparison.OrdinalIgnoreCase))
                {
                    _powerPinIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderToggleDirection, StringComparison.OrdinalIgnoreCase))
                {
                    _toggleDirectionIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderToggleThresholdFailCount, StringComparison.OrdinalIgnoreCase))
                {
                    _toggleThresholdFailCountIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderAnalogSweepPin, StringComparison.OrdinalIgnoreCase))
                {
                    _analogSweepPinIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderAnalogSweepStart, StringComparison.OrdinalIgnoreCase))
                {
                    _analogSweepStartIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderAnalogSweepStop, StringComparison.OrdinalIgnoreCase))
                {
                    _analogSweepStopIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderAnalogSweepStep, StringComparison.OrdinalIgnoreCase))
                {
                    _analogSweepStepIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimLinkSweepStart, StringComparison.OrdinalIgnoreCase))
                {
                    _trimLinkSweepStartIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTrimLinkSweepStop, StringComparison.OrdinalIgnoreCase))
                {
                    _trimLinkSweepStopIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderGngLow, StringComparison.OrdinalIgnoreCase))
                {
                    _gNgLowIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderGngHigh, StringComparison.OrdinalIgnoreCase))
                {
                    _gNgHighIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderLowLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _lowLimitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderHighLimit, StringComparison.OrdinalIgnoreCase)) _highLimitIndex = i;
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
                        .Equals(ConHeaderFunctionName, StringComparison.OrdinalIgnoreCase))
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
            _functionNameIndex = -1;
            _blockNameIndex = -1;
            _trimMethodIndex = -1;
            _trimTableIndex = -1;
            _codeDistributionIndex = -1;
            _numBitsIndex = -1;
            _lSbIndex = -1;
            _targetIndex = -1;
            _preTrimCodeIndex = -1;
            _trimRegisterIndex = -1;
            _trimBitFieldIndex = -1;
            _oTpRegisterIndex = -1;
            _measPinIndex = -1;
            _waitTimeIndex = -1;
            _sampleSizeIndex = -1;
            _datalogTemplateIndex = -1;
            _powerPinIndex = -1;
            _toggleDirectionIndex = -1;
            _toggleThresholdFailCountIndex = -1;
            _analogSweepPinIndex = -1;
            _analogSweepStartIndex = -1;
            _analogSweepStopIndex = -1;
            _analogSweepStepIndex = -1;
            _trimLinkSweepStartIndex = -1;
            _trimLinkSweepStopIndex = -1;
            _gNgLowIndex = -1;
            _gNgHighIndex = -1;
            _lowLimitIndex = -1;
            _highLimitIndex = -1;
        }
    }

    internal class InstanceRowComparer : IEqualityComparer<InstanceRow>
    {
        public bool Equals(InstanceRow x, InstanceRow y)
        {
            if (x == null || y == null)
                return false;
            if (x.TestName == null || y.TestName == null)
                return false;

            if (x.TestName.Equals(y.TestName, StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        public int GetHashCode(InstanceRow obj)
        {
            if (obj == null)
                return 0;
            if (obj.TestName == null)
                return 0;

            return obj.TestName.ToLower().GetHashCode();
        }
    }

    internal class FlowRowComparer : IEqualityComparer<FlowRow>
    {
        public bool Equals(FlowRow x, FlowRow y)
        {
            if (x == null || y == null)
                return false;
            if (x.Parameter == null || y.Parameter == null)
                return false;

            if (x.Parameter.Equals(y.Parameter, StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        public int GetHashCode(FlowRow obj)
        {
            if (obj == null)
                return 0;
            if (obj.Parameter == null)
                return 0;

            return obj.Parameter.ToLower().GetHashCode();
        }
    }
}