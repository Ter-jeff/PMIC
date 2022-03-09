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
// 2021-12-23  Bruce          #273	          PMICToolBox_v2021.12.17.1, VBTPOP Gen PreCheck function enhancement
//
//------------------------------------------------------------------------------ 
using OfficeOpenXml;
using OfficeOpenXml.Style;
using VBTPOPGenPreCheckBusiness.DataStore;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Diagnostics;

namespace VBTPOPGenPreCheckBusiness.Business
{
    public class CheckProcess
    {
        private GuiInfo _guiInfo = null;

        private List<string> _testplanList = null;
        private string _otpRegMap;
        private string _ahbRegMap;
        private string _pinMap;
        private string _output;


        public CheckProcess(GuiInfo guiInfo)
        {
            _guiInfo = guiInfo;
            _testplanList = guiInfo.testplanList;
            _otpRegMap = guiInfo.otpRegMap;
            _ahbRegMap = guiInfo.AhbRegMap;
            _pinMap = guiInfo.PinMap;
            _output = guiInfo.output;
        }

        public CheckProcess(string testplanPath, string otpRegMapPath, string ahbRegMapPath, string pinMapPath, string outputPath)
        {
            _testplanList = new List<string>() { testplanPath };
            _otpRegMap = otpRegMapPath;
            _ahbRegMap = ahbRegMapPath;
            _pinMap = pinMapPath;
            _output = outputPath;
        }

        private void UpdateGuiStatus(string progressValue, string uiSateInfo)
        {
            if (_guiInfo != null)
            {
                _guiInfo.ProgressValue = progressValue;
                _guiInfo.UISateInfo = uiSateInfo;
            }
        }
        public void WorkFlow()
        {
            try
            {
                List<TestPlanFile> listTestplanData = new List<TestPlanFile>();

                UpdateGuiStatus("5", "Reading TestPlan files...");

                // multiple testplan reader
                foreach (string tp in _testplanList)
                {
                    if (string.IsNullOrEmpty(tp) || !File.Exists(tp))
                        continue;
                    var tpReader = new TestplanReader(tp);
                    if (tpReader.IsValidTestplan)
                    {
                        TestPlanFile testPlan = tpReader.Read();
                        if (testPlan != null)
                            listTestplanData.Add(testPlan);
                    }
                }
                if (listTestplanData.Count == 0)
                    throw new Exception("No valid testPlan exists!");

                OTPRegisterMapReader otpRegisterMap = null;
                AhbRegisterMapSheet ahbRegisterMap = null;
                PinMapSheet pinMap = null;
                // OTPRegisterMap reader
                if (!string.IsNullOrEmpty(_otpRegMap) && File.Exists(_otpRegMap))
                {
                    UpdateGuiStatus("45", "Reading OTP Register Map File...");
                    otpRegisterMap = new OTPRegisterMapReader(_otpRegMap);
                }
                //AHBRegisterMap
                if (!string.IsNullOrEmpty(_ahbRegMap) && File.Exists(_ahbRegMap))
                {
                    UpdateGuiStatus("55", "Reading AHB Register Map File...");
                    ahbRegisterMap = new AhbRegisterMapReader().Read(_ahbRegMap);
                }
                //PinMap
                if (!string.IsNullOrEmpty(_pinMap) && File.Exists(_pinMap))
                {
                    UpdateGuiStatus("65", "Reading Pin Map File...");
                    pinMap = new PinMapReader().ReadSheet(_pinMap);
                }

                // gen result
                string outputReport = Path.Combine(_output, "OTP_Check_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                GenReport(listTestplanData, otpRegisterMap.dicOTPRegMap, ahbRegisterMap, pinMap, outputReport);
                UpdateGuiStatus("100", "Complete!");

                if (File.Exists(outputReport))
                    Process.Start(outputReport);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void GenReport(List<TestPlanFile> listTestplanData, Dictionary<string, OTPRegisterMapData> dicOTPRegMap, AhbRegisterMapSheet ahbRegisterMap,
                                PinMapSheet pinMap, string outputReport)
        {
            using (ExcelPackage ep = new ExcelPackage(new FileInfo(outputReport)))
            {
                if (dicOTPRegMap != null)
                {
                    UpdateGuiStatus("75", "Checking OTP Owner...");
                    GenOTPOwnerResult(ep, listTestplanData, dicOTPRegMap);

                    UpdateGuiStatus("80", "Checking Test Paremeter Trim Register...");
                    GenTrimRegisterResult(ep, listTestplanData, dicOTPRegMap);
                }

                if (ahbRegisterMap != null)
                {
                    UpdateGuiStatus("85", "Checking Test Flow Register...");
                    GenTestFlowRegisterCheckResult(ep, listTestplanData, ahbRegisterMap, dicOTPRegMap);
                }

                if (pinMap != null)
                {
                    UpdateGuiStatus("95", "Checking TestPlan Pins...");
                    GenPinCheckResult(ep, listTestplanData, pinMap);
                }

                ep.Save();
            }
        }

        /// <summary>
        /// Check Test Paremeter TrimRegister/TrimBitField/Numbits with OTPRegisterMap
        /// </summary>
        /// <param name="ep"></param>
        /// <param name="listTestplanData"></param>
        /// <param name="dicOTPRegMap"></param>
        private void GenTrimRegisterResult(ExcelPackage ep, List<TestPlanFile> listTestplanData, Dictionary<string, OTPRegisterMapData> dicOTPRegMap)
        {
            ExcelWorksheet ws = ep.Workbook.Worksheets.Add("Check_TrimRegister_Result");

            // header line
            ws.Cells[1, 1].Value = "Testplan Filename";
            ws.Cells[1, 2].Value = "OTP_REGISTER_NAME";
            ws.Cells[1, 3].Value = "TrimRegister"; // compare with reg_name
            ws.Cells[1, 4].Value = "TrimBitField"; // compare with name
            ws.Cells[1, 5].Value = "Numbits"; // compare with bw
            ws.Cells[1, 6].Value = "Result";

            for (int c = 1; c <= 6; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            int rowIndex = 2;
            foreach (TestPlanFile testPlan in listTestplanData)
            {

                ws.Cells[rowIndex, 1].Value = testPlan.FileName;
                foreach (var item in testPlan.ParameterSheet.TestParameterRows)
                {
                    if (isNA(item.TrimRegister))
                        continue;
                    string otpRegister = item.OtpRegister;
                    string trimRegister = item.TrimRegister;
                    string trimBitField = item.TrimBitField;
                    string numbits = item.Numbits;
                    ws.Cells[rowIndex, 2].Value = otpRegister;
                    ws.Cells[rowIndex, 3].Value = trimRegister;
                    ws.Cells[rowIndex, 4].Value = trimBitField;
                    ws.Cells[rowIndex, 5].Value = numbits;

                    if (dicOTPRegMap.ContainsKey(otpRegister))
                    {
                        var map = dicOTPRegMap[otpRegister];
                        if (map.regName.ToUpper().Equals(trimRegister.ToUpper()) && map.name.ToUpper().Equals(trimBitField.ToUpper()) && map.bw.Equals(numbits))
                            ws.Cells[rowIndex, 6].Value = "V";
                        else
                        {
                            if (item.MultiBlock)
                            {
                                ws.Cells[rowIndex, 3].Value = string.Empty;
                                ws.Cells[rowIndex, 4].Value = string.Empty;
                                ws.Cells[rowIndex, 5].Value = string.Empty;
                                ws.Cells[rowIndex, 6].Value = "N/A";
                            }
                            else
                            {
                                ws.Cells[rowIndex, 6].Value = "X";
                            }
                            ws.Cells[rowIndex, 6].Style.Font.Color.SetColor(Color.Red);
                            ws.Cells[rowIndex, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[rowIndex, 6].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                            if (!map.regName.ToUpper().Equals(trimRegister.ToUpper()))
                            {
                                ws.Cells[rowIndex, 3].Style.Font.Color.SetColor(Color.Red);
                                ws.Cells[rowIndex, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[rowIndex, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                ws.Cells[rowIndex, 3].AddComment(map.regName, "Teradyne");
                            }

                            if (!map.name.ToUpper().Equals(trimBitField.ToUpper()))
                            {
                                ws.Cells[rowIndex, 4].Style.Font.Color.SetColor(Color.Red);
                                ws.Cells[rowIndex, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[rowIndex, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                ws.Cells[rowIndex, 4].AddComment(map.name, "Teradyne");
                            }

                            if (!map.bw.Equals(numbits))
                            {
                                ws.Cells[rowIndex, 5].Style.Font.Color.SetColor(Color.Red);
                                ws.Cells[rowIndex, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[rowIndex, 5].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                ws.Cells[rowIndex, 5].AddComment(map.bw, "Teradyne");
                            }
                        }
                    }
                    else
                    {
                        if (item.MultiBlock)
                        {
                            ws.Cells[rowIndex, 3].Value = string.Empty;
                            ws.Cells[rowIndex, 4].Value = string.Empty;
                            ws.Cells[rowIndex, 5].Value = string.Empty;
                        }
                        ws.Cells[rowIndex, 6].Value = "X";
                        for (int c = 2; c <= 6; ++c)
                        {
                            ws.Cells[rowIndex, c].Style.Font.Color.SetColor(Color.Red);
                            ws.Cells[rowIndex, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[rowIndex, c].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        }
                    }
                    ++rowIndex;
                }
            }

            ws.Cells.AutoFitColumns();
        }

        /// <summary>
        /// 1.Check parameter sheet TrimRegister column
        /// if TrimRegister=Empty||NA||N/A ignore
        /// if TrimRegister exist in OTPRegisterMap and owner != trim, fail, and record owner to report
        /// if TrimRegister not exist in OTPRegisterMap fail, and record NA to report
        /// 
        /// 2.List all the unused(by parameter OtpRegister) trim register name(register name which owner is trim) in OTP Register Map
        /// </summary>
        /// <param name="ep"></param>
        /// <param name="listTestplanData"></param>
        /// <param name="dicOTPRegMap"></param>
        private void GenOTPOwnerResult(ExcelPackage ep, List<TestPlanFile> listTestplanData, Dictionary<string, OTPRegisterMapData> dicOTPRegMap)
        {
            ExcelWorksheet ws = ep.Workbook.Worksheets.Add("Check_OTPOwner_Result");

            // header line
            ws.Cells[1, 1].Value = "Testplan Filename";
            ws.Cells[1, 2].Value = "Sheet Name";
            ws.Cells[1, 3].Value = "OTP_REGISTER_NAME";
            ws.Cells[1, 4].Value = "otp_owner";
            ws.Cells[1, 5].Value = "Result";

            for (int c = 1; c <= 5; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            int rowIndex = 2;
            foreach (TestPlanFile testPlan in listTestplanData)
            {
                int statusIndex = rowIndex; // used to keep pass/fail row index
                ws.Cells[rowIndex, 1].Value = testPlan.FileName;

                int failCount = 0;
                foreach (var item in testPlan.ParameterSheet.TestParameterRows)
                {
                    if (isNA(item.TrimRegister))
                        continue;
                    if (dicOTPRegMap.ContainsKey(item.OtpRegister))
                    {
                        string otp_owner = dicOTPRegMap[item.OtpRegister].otp_owner;
                        if (!otp_owner.ToLower().Equals("trim"))
                        {
                            ws.Cells[rowIndex, 2].Value = testPlan.ParameterSheet.SheetName;
                            ws.Cells[rowIndex, 3].Value = item.OtpRegister;
                            ws.Cells[rowIndex, 4].Value = otp_owner;
                            ++rowIndex;
                            ++failCount;
                        }
                    }
                    else
                    {
                        ws.Cells[rowIndex, 2].Value = testPlan.ParameterSheet.SheetName;
                        ws.Cells[rowIndex, 3].Value = item.OtpRegister;
                        ws.Cells[rowIndex, 4].Value = "N/A";
                        ++rowIndex;
                        ++failCount;
                    }
                }

                foreach (var testFlowSheet in testPlan.TestFlowSheetlst)
                {
                    foreach (var item in testFlowSheet.CommandRows)
                    {
                        if (!isOtpWrite(item.CommandName))
                            continue;
                        if (isNA(item.RegisterName))
                            continue;
                        if (dicOTPRegMap.ContainsKey(item.RegisterName))
                        {
                            string otp_owner = dicOTPRegMap[item.RegisterName].otp_owner;
                            if (!otp_owner.ToLower().Equals("trim"))
                            {
                                ws.Cells[rowIndex, 2].Value = testFlowSheet.SheetName;
                                ws.Cells[rowIndex, 3].Value = item.RegisterName;
                                ws.Cells[rowIndex, 4].Value = otp_owner;
                                ++rowIndex;
                                ++failCount;
                            }
                        }
                        else
                        {
                            ws.Cells[rowIndex, 2].Value = testFlowSheet.SheetName;
                            ws.Cells[rowIndex, 3].Value = item.RegisterName;
                            ws.Cells[rowIndex, 4].Value = "N/A";
                            ++rowIndex;
                            ++failCount;
                        }
                    }
                }

                if (failCount != 0)
                {
                    ws.Cells[statusIndex, 5].Value = "Fail";
                    ws.Cells[statusIndex, 5].Style.Font.Color.SetColor(Color.Red);
                }
                else
                {
                    ws.Cells[statusIndex, 5].Value = "Pass";
                    ws.Cells[statusIndex, 5].Style.Font.Color.SetColor(Color.Green);
                }

                ++rowIndex;
            }

            ws.Cells.AutoFitColumns();

            ws = ep.Workbook.Worksheets.Add("Check_UnUsedRegister_Result");
            ws.Cells[1, 1].Value = "UnUsed Register Names(Trim)";
            for (int c = 1; c <= 1; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            rowIndex = 2;
            foreach (string register in GetUnusedTrimRegister(listTestplanData, dicOTPRegMap))
            {
                ws.Cells[rowIndex, 1].Value = register;
                ++rowIndex;
            }
            ws.Cells.AutoFitColumns();
        }


        private List<string> GetUnusedTrimRegister(List<TestPlanFile> listTestplanData, Dictionary<string, OTPRegisterMapData> dicOTPRegMap)
        {
            List<string> unusedTrimRegisterlst = new List<string>();
            List<string> usedRegisterlst = new List<string>();
            List<string> trimRegisterlstInRegMap = dicOTPRegMap.Keys.ToList().FindAll(s => dicOTPRegMap[s].otp_owner.Equals("trim", StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (TestPlanFile testPlan in listTestplanData)
            {
                foreach (TestParameterRow parameterRow in testPlan.ParameterSheet.TestParameterRows)
                {
                    usedRegisterlst.Add(parameterRow.OtpRegister.ToUpper());
                }
                // 2021-12-23  Bruce          #273	          PMICToolBox_v2021.12.17.1, VBTPOP Gen PreCheck function enhancement add start
                foreach (var testFlowSheet in testPlan.TestFlowSheetlst)
                {
                    foreach (var commandRow in testFlowSheet.CommandRows)
                    {
                        if (commandRow.TopList.Equals("OTP",StringComparison.InvariantCultureIgnoreCase) &&
                            commandRow.CommandName.Equals("OTP_WRITE", StringComparison.InvariantCultureIgnoreCase))
                        {
                            usedRegisterlst.Add(commandRow.RegisterName.ToUpper());
                        }
                    }
                }
                // 2021-12-23  Bruce          #273	          PMICToolBox_v2021.12.17.1, VBTPOP Gen PreCheck function enhancement add end
            }
            usedRegisterlst = usedRegisterlst.Distinct().ToList();
            unusedTrimRegisterlst = trimRegisterlstInRegMap.FindAll(s => !usedRegisterlst.Exists(m => m.Equals(s, StringComparison.OrdinalIgnoreCase)));
            return unusedTrimRegisterlst;
        }

        /// <summary>
        /// Check testPlan sheet(function must define in parameter sheet) commands reg_name/bf_name/values with AhbRegisterMap reg name/field name/field width
        /// only check "OTP_WRITE", "AHB_READ", "AHB_WRITE", "VM_WRITE" commands
        /// for OTP_WRITE only check testplan reg_name must not be null and must exist in OTPRegisterMap
        /// for others commands:
        /// 1.check  testplan reg_name must not be null
        /// 2.testplan reg_name must exist in AHBRegisterMap
        /// 3.if testplan bf_name is not null, bf_name must exist in AHBRegisterMap
        /// 4.find AHBRegisterMap row by reg_name and bf_name, and get field width, check testplan values must meet with field width
        /// eg:field width: 1  values must be &H0 ~&HFF
        /// </summary>
        /// <param name="ep"></param>
        /// <param name="listTestplanData"></param>
        /// <param name="ahbRegMapSheet"></param>
        /// <param name="dicOTPRegMap"></param>
        private void GenTestFlowRegisterCheckResult(ExcelPackage ep, List<TestPlanFile> listTestplanData, AhbRegisterMapSheet ahbRegMapSheet, Dictionary<string, OTPRegisterMapData> dicOTPRegMap)
        {
            ExcelWorksheet ws = ep.Workbook.Worksheets.Add("Check_TestFlow_Register_Result");
            // header line
            ws.Cells[1, 1].Value = "Testplan Filename";
            ws.Cells[1, 2].Value = "Sheet Name";
            ws.Cells[1, 3].Value = "Row";
            ws.Cells[1, 4].Value = "Command";
            ws.Cells[1, 5].Value = "REGISTER";
            ws.Cells[1, 6].Value = "BITFIELD NAME";
            ws.Cells[1, 7].Value = "VALUE(S)";
            ws.Cells[1, 8].Value = "Result";

            string bitFieldName;

            for (int c = 1; c <= 8; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            int rowIndex = 2;
            foreach (TestPlanFile testPlan in listTestplanData)
            {
                ws.Cells[rowIndex, 1].Value = testPlan.FileName;
                foreach (TestFlowSheet testFlowSheet in testPlan.FindValidTestFlowSheets())
                {
                    ws.Cells[rowIndex, 2].Value = testFlowSheet.SheetName;
                    rowIndex++;
                    foreach (CommandRow command in testFlowSheet.GetCommandlistToCheckRegister())
                    {
                        ws.Cells[rowIndex, 3].Value = command.RowIndex.ToString();
                        ws.Cells[rowIndex, 4].Value = command.CommandName;
                        ws.Cells[rowIndex, 5].Value = command.RegisterName;
                        ws.Cells[rowIndex, 6].Value = command.BitFieldName;
                        //Request from alec, AHB_READ no need to fill VALUES in testplan
                        if (!command.CommandName.Equals("AHB_READ"))
                            ws.Cells[rowIndex, 7].Value = command.Values;

                        string result = "V";
                        bool seprateFlag = false;
                        if (string.IsNullOrEmpty(command.RegisterName))
                        {
                            result = "Missing Register Name in TestFlow";
                            FormatErrorCell(ws, rowIndex, 5);
                        }
                        else
                        {

                            if (command.CommandName.Equals("OTP_WRITE", StringComparison.OrdinalIgnoreCase))
                            {
                                if (dicOTPRegMap != null && !dicOTPRegMap.Keys.ToList().Exists(s => s.Equals(command.RegisterName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    result = "Can not find OTP Register Name in OTP_Register_Map";
                                    FormatErrorCell(ws, rowIndex, 5);
                                }
                            }
                            else
                            {
                                var ahbRegRows = ahbRegMapSheet.AhbRegRows.FindAll(s => s.RegName.Equals(command.RegisterName, StringComparison.OrdinalIgnoreCase));
                                if (ahbRegRows.Count == 0 && command.RegisterName.Contains("."))
                                {
                                    ahbRegRows = ahbRegMapSheet.AhbRegRows.FindAll(s => s.RegName.Equals(command.RegisterName.Split('.')[0].Trim(), StringComparison.OrdinalIgnoreCase));
                                    seprateFlag = true;
                                }

                                if (ahbRegRows == null || ahbRegRows.Count == 0)
                                {
                                    result = "Can not find Register Name in AHB_Register_Map";
                                    FormatErrorCell(ws, rowIndex, 5);
                                }
                                else
                                {
                                    bitFieldName = command.BitFieldName.Trim();
                                    if (string.IsNullOrEmpty(bitFieldName) && seprateFlag)
                                        bitFieldName = command.RegisterName.Split('.')[1].Trim();
                                    var ahbRegRow = ahbRegRows.Find(s => s.FieldName.Equals(bitFieldName, StringComparison.OrdinalIgnoreCase));
                                    if (ahbRegRow != null)
                                    {
                                        //Request from alec, AHB_READ no need to check VALUES in testplan
                                        if (!command.CommandName.Equals("AHB_READ"))
                                        {
                                            int fieldWidth = 0;
                                            int.TryParse(ahbRegRow.FieldWidth, out fieldWidth);
                                            if (fieldWidth > 0 && !string.IsNullOrEmpty(command.Values))
                                            {
                                                int minValue = 0;
                                                int maxValue = (int)Math.Pow(2, (double)fieldWidth) - 1;
                                                int iValue = -1;
                                                bool formatPass = int.TryParse(command.Values.ToUpper().Replace("&H", "").Replace("0X", ""), System.Globalization.NumberStyles.HexNumber, null, out iValue);
                                                if (!formatPass || iValue < minValue || iValue > maxValue)
                                                {
                                                    result = !formatPass ? "Value(s) Type Mismatch" : "Value(s) is not match with Filed Width";
                                                    ws.Cells[rowIndex, 7].AddComment(string.Format("Should be &H0 ~ {0}", "&H" + String.Format("{0:X}", maxValue)), "Teradyne");
                                                    FormatErrorCell(ws, rowIndex, 7);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(bitFieldName))
                                            result = "Can not find BitFildName in AHB_Register_Map";
                                    }

                                    if (seprateFlag && result == "V")
                                        result = "V(should separate the register by register and bit field)";
                                }
                            }
                        }

                        ws.Cells[rowIndex, 8].Value = result;

                        if (result != "V")
                        {
                            FormatErrorCell(ws, rowIndex, 8);
                        }

                        ++rowIndex;
                    }
                }
            }

            ws.Cells.AutoFitColumns();
        }

        /// <summary>
        /// Check testplan(test flow and parameter) pin missing in pinmap
        /// </summary>
        /// <param name="ep"></param>
        /// <param name="listTestplanData"></param>
        /// <param name="pinMapSheet"></param>
        private void GenPinCheckResult(ExcelPackage ep, List<TestPlanFile> listTestplanData, PinMapSheet pinMapSheet)
        {
            ExcelWorksheet ws = ep.Workbook.Worksheets.Add("Check_TestFlow_Pin_Result");

            // header line
            ws.Cells[1, 1].Value = "Testplan Filename";
            ws.Cells[1, 2].Value = "Sheet Name";
            ws.Cells[1, 3].Value = "Row";
            ws.Cells[1, 4].Value = "Command";
            ws.Cells[1, 5].Value = "PIN";
            ws.Cells[1, 6].Value = "Result";

            for (int c = 1; c <= 6; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            int rowIndex = 2;
            foreach (TestPlanFile testPlan in listTestplanData)
            {
                ws.Cells[rowIndex, 1].Value = testPlan.FileName;
                foreach (TestFlowSheet testFlowSheet in testPlan.FindValidTestFlowSheets())
                {
                    ws.Cells[rowIndex, 2].Value = testFlowSheet.SheetName;
                    rowIndex++;
                    foreach (CommandRow command in testFlowSheet.GetCommandlistToCheckPin())
                    {
                        ws.Cells[rowIndex, 3].Value = command.RowIndex.ToString();
                        ws.Cells[rowIndex, 4].Value = command.CommandName;
                        ws.Cells[rowIndex, 5].Value = command.Pin;

                        string result = "V";
                        string missingPin = "";
                        if (!CheckPins(command.Pin, pinMapSheet, out missingPin))
                        {
                            result = "X";
                            ws.Cells[rowIndex, 5].AddComment(string.Format("Missing pins: {0} in pinMap", missingPin), "Teradyne");
                        }
                        ws.Cells[rowIndex, 6].Value = result;
                        if (result != "V")
                        {
                            FormatErrorCell(ws, rowIndex, 5);
                            FormatErrorCell(ws, rowIndex, 6);
                        }
                        ++rowIndex;
                    }
                }
            }

            ws.Cells.AutoFitColumns();


            ws = ep.Workbook.Worksheets.Add("Check_Parameter_Pin_Result");
            // header line
            ws.Cells[1, 1].Value = "Testplan Filename";
            ws.Cells[1, 2].Value = "Row";
            ws.Cells[1, 3].Value = "FunctionName";
            ws.Cells[1, 4].Value = "BlockName";
            ws.Cells[1, 5].Value = "MeasPin";
            ws.Cells[1, 6].Value = "PowerPin";
            ws.Cells[1, 7].Value = "AnalogPin";
            ws.Cells[1, 8].Value = "Result";

            for (int c = 1; c <= 8; ++c)
            {
                ws.Cells[1, c].Style.Font.Bold = true;
                ws.Cells[1, c].Style.Font.Color.SetColor(Color.White);
                ws.Cells[1, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, c].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
            }

            rowIndex = 2;
            foreach (TestPlanFile testPlan in listTestplanData)
            {
                ws.Cells[rowIndex, 1].Value = testPlan.FileName;

                foreach (TestParameterRow parameterRow in testPlan.ParameterSheet.TestParameterRows)
                {
                    ws.Cells[rowIndex, 2].Value = parameterRow.Row.ToString();
                    ws.Cells[rowIndex, 3].Value = parameterRow.FunctionName;
                    ws.Cells[rowIndex, 4].Value = parameterRow.BlockName;
                    ws.Cells[rowIndex, 5].Value = parameterRow.MeasPin;
                    ws.Cells[rowIndex, 6].Value = parameterRow.PowerPin;
                    ws.Cells[rowIndex, 7].Value = parameterRow.AnalogPin;
                    string result = "V";
                    string missingPin = "";
                    if (!CheckPins(parameterRow.MeasPin, pinMapSheet, out missingPin))
                    {
                        result = "X";
                        ws.Cells[rowIndex, 5].AddComment(string.Format("Missing pins: {0} in pinMap", missingPin), "Teradyne");
                        FormatErrorCell(ws, rowIndex, 5);
                    }
                    if (!CheckPins(parameterRow.PowerPin, pinMapSheet, out missingPin))
                    {
                        result = "X";
                        ws.Cells[rowIndex, 6].AddComment(string.Format("Missing pins: {0} in pinMap", missingPin), "Teradyne");
                        FormatErrorCell(ws, rowIndex, 6);
                    }
                    if (!CheckPins(parameterRow.AnalogPin, pinMapSheet, out missingPin))
                    {
                        result = "X";
                        ws.Cells[rowIndex, 7].AddComment(string.Format("Missing pins: {0} in pinMap", missingPin), "Teradyne");
                        FormatErrorCell(ws, rowIndex, 7);
                    }
                    ws.Cells[rowIndex, 8].Value = result;
                    if (result != "V")
                    {
                        FormatErrorCell(ws, rowIndex, 8);
                    }
                    ++rowIndex;
                }

            }
            ws.Cells.AutoFitColumns();
        }

        private void FormatErrorCell(ExcelWorksheet ws, int row, int col)
        {
            ws.Cells[row, col].Style.Font.Color.SetColor(Color.Red);
            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }

        //10/26/2021 update, support pin format "pin1$pin2", symbol “$” available
        private bool CheckPins(string pin, PinMapSheet pinMapSheet, out string missingPins)
        {
            missingPins = "";
            if (string.IsNullOrEmpty(pin))
                return true;
            List<string> pinlst = pin.Split(',', '$').ToList();
            pinlst = pinlst.Select(s => s.Trim()).ToList();
            pinlst.RemoveAll(s => s.Equals("N/A", StringComparison.OrdinalIgnoreCase) || s.Equals("NA", StringComparison.OrdinalIgnoreCase));
            List<string> missingPinlst = pinlst.FindAll(p => !pinMapSheet.PinList.Exists(s => s.pinName.Equals(p, StringComparison.OrdinalIgnoreCase)) &&
                                                           !pinMapSheet.PinGroupList.Exists(m => m.GroupName.Equals(p, StringComparison.OrdinalIgnoreCase)));
            if (missingPinlst != null && missingPinlst.Count > 0)
            {
                missingPins = string.Join(",", missingPinlst);
                return false;
            }
            return true;
        }

        private bool isNA(string input)
        {
            input = input.ToUpper().Trim();
            if (input.Equals("N/A") || input.Equals("NA") || string.IsNullOrEmpty(input))
                return true;
            return false;
        }

        private bool isOtpWrite(string input)
        {
            input = input.ToUpper().Trim();
            if (input.Equals("OTP_WRITE"))
                return true;
            return false;
        }
    }
}
