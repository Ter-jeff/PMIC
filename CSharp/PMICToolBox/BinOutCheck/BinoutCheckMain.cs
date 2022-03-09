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
// 2022-02-21  Steven Chen    #320            Support No FlowHeader in Binoutcheck
// 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut
// 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item
// 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet.
// 2021-12-15  Bruce          #265	          Commend Flow
// 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify
// 2021-12-15  Bruce          #263	          Error Message Modify: Fail Flag count is not match with item‘s count ->FlagDefaultValue Error
// 2021-12-15  Bruce          #262	          Set New Color and adjust Color area
// 2021-12-15  Bruce          #261	          New Column For Fail Flag-Symbol Define Of Fail Flag
// 2021-12-08  Bruce          #258	          Add No Bin Out Check
// 2021-12-08  Bruce          #256	          Check Flag Default Value
// 2021-12-05  Bruce          #255	          Check Fail flag and BinTable Sequence Check by fail flag
// 2021-12-05  Bruce          #254	          Missing Color in column SwBinDuplicate
// 2021-12-05  Bruce          #247	          Fail_Stop_Table.txt support muti column and blanks on the bottom
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Windows.Forms;
using IgxlData.IgxlBase;

namespace BinOutCheck
{
    public class BinoutCheckMain
    {
        private const string BinTable = "Bintable";
        private const string Test = "Test";
        private const string SetErrorBin = "set-error-bin";
        // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add start
        private readonly Color ErrorColor = Color.FromArgb(248, 203, 173);
        // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add end
        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
        private readonly Color ErrorColor2 = Color.FromArgb(255, 255, 0);
        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
        private readonly List<BinOutConfig> _binoutCfg;
        private readonly List<BinOutCheckResult> _results;
        private readonly List<ErrorResult> _errorResults;
        private readonly List<SetErrorBinResult> _setErrorBinResults;
        private readonly List<FlowSetBinOutReport> _functionBinResults;
        private readonly List<VBTSetBinCheckResult> _VBTSetBinCheckResults;
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
        private readonly List<NoBinOutResult> _NoBinOutResultResults;
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
        private readonly List<NoBinOutReferenceDatalogResult> _NoBinOutReferenceDatalogResults;
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

        private IgxlProgram _testProgram;
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut del start
        //private string _outputFilePath;
        //private string _tpFilePath;
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut del end
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
        private StringIgnoreCaseCompare _StringIgnoreCaseCompare;
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
        private GuiArgs _guiArgs;
        private DataLogFileInfo _DataLogFileInfo = new DataLogFileInfo();
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
        //public BinoutCheckMain(string tpPath, string outputFilePath = null)
        public BinoutCheckMain(GuiArgs guiArgs)
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
        {
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
            //_tpFilePath = tpPath;
            _guiArgs = guiArgs;
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
            //if (outputFilePath != null)
            //{
            //    _outputFilePath = outputFilePath;
            //    if (!Directory.Exists(new FileInfo(_outputFilePath).DirectoryName))
            //        throw new Exception("Output file parent folder is not exist: " + new FileInfo(_outputFilePath).DirectoryName);
            //    if (!_outputFilePath.EndsWith(".xlsx"))
            //        throw new Exception("Output file is not an xlsx excel file: " + _outputFilePath);
            //}
            //else
            //{
            //    FileInfo tpFileInfo = new FileInfo(tpPath);
            //    _outputFilePath = Path.Combine(tpFileInfo.DirectoryName, Regex.Replace(tpFileInfo.Name, @"\..+", "_BinOutCheck_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx"));
            //}

            if (_guiArgs.OutputFilePath != null)
            {
                if (!Directory.Exists(new FileInfo(_guiArgs.OutputFilePath).DirectoryName))
                    throw new Exception("Output file parent folder is not exist: " + new FileInfo(_guiArgs.OutputFilePath).DirectoryName);
                if (!_guiArgs.OutputFilePath.EndsWith(".xlsx"))
                    throw new Exception("Output file is not an xlsx excel file: " + _guiArgs.OutputFilePath);
            }
            else
            {
                FileInfo tpFileInfo = new FileInfo(_guiArgs.TestPlanPath);
                _guiArgs.OutputFilePath = Path.Combine(tpFileInfo.DirectoryName, Regex.Replace(tpFileInfo.Name, @"\..+", "_BinOutCheck_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx"));
            }

            if (!string.IsNullOrEmpty(_guiArgs.TestPlanPath))
            {
                if (!File.Exists(_guiArgs.TestPlanPath))
                    throw new Exception("TestPlan file is not exist: " + _guiArgs.TestPlanPath);
            }

            if (!string.IsNullOrEmpty(_guiArgs.DataLogPath))
            {
                string errorMsg = _DataLogFileInfo.PreCheckDataLog(_guiArgs.DataLogPath);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    throw new Exception(errorMsg);
                }
            }
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end

            _results = new List<BinOutCheckResult>();
            _binoutCfg = new List<BinOutConfig>();
            _errorResults = new List<ErrorResult>();
            _setErrorBinResults = new List<SetErrorBinResult>();
            _functionBinResults = new List<FlowSetBinOutReport>();
            _VBTSetBinCheckResults = new List<VBTSetBinCheckResult>();
            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
            _NoBinOutResultResults = new List<NoBinOutResult>();
            _StringIgnoreCaseCompare = new StringIgnoreCaseCompare();
            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
            _NoBinOutReferenceDatalogResults = new List<NoBinOutReferenceDatalogResult>();
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end
        }
        //public BinoutCheckMain(IgxlProgram tp, string outputFolder = null)
        //{
        //    _testProgram = tp;
        //    _results = new List<BinOutCheckResult>();
        //    _errorResults = new List<ErrorResult>();
        //    _setErrorBinResults = new List<SetErrorBinResult>();
        //    _functionBinResults = new List<FlowSetBinOutReport>();
        //    _VBTSetBinCheckResults = new List<VBTSetBinCheckResult>();
        //    if (outputFolder != null)
        //        _outputFolder = outputFolder;
        //    else
        //    {
        //        _outputFolder = new FileInfo(_testProgram.TpPath).DirectoryName;
        //    }

        //    CreateResultFrame();
        //}

        private void CreateResultFrame()
        {
            foreach (var sheet in _testProgram.BintableSheets)
            {
                foreach (var row in sheet.BinTableRows)
                {
                    var binOutCheckResult = new BinOutCheckResult();
                    binOutCheckResult.Name = row.Name;
                    binOutCheckResult.ItemList = row.ItemList;
                    binOutCheckResult.Op = row.Op;
                    binOutCheckResult.Sort = row.Sort;
                    binOutCheckResult.Bin = row.Bin;
                    binOutCheckResult.Result = row.Result;
                    _results.Add(binOutCheckResult);
                }
            }
        }

        private void LoadConfig()
        {
            string cfgFile = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "Config", "Fail_Stop_Table.txt");
            string[] allLines = File.ReadAllLines(cfgFile);
            // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add start
            List<string> items;
            // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add end
            foreach (var line in allLines)
            {
                // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add start
                items = GetItems(line);
                if (items.Count >= 6)
                {
                    //2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add end
                    BinOutConfig binoutcfg = new BinOutConfig();
                    binoutcfg.Name = items[0];
                    binoutcfg.ItemList = items[1];
                    binoutcfg.Op = items[2];
                    binoutcfg.Sort = items[3];
                    binoutcfg.Bin = items[4];
                    binoutcfg.Result = items[5];
                    _binoutCfg.Add(binoutcfg);
                    // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add start
                }
                // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add end
            }
        }

        // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item add start
        private void PreProcessData(IgxlProgram program)
        {
            // Clear Dummy Data
            foreach (var flowSheet in program.FlowSheets)
            {
                int blankRowIndex = -1;
                for (int i = 0; i < flowSheet.FlowRows.Count; i++)
                {
                    if (string.IsNullOrEmpty(flowSheet.FlowRows[i].Opcode) && string.IsNullOrEmpty(flowSheet.FlowRows[i].Parameter))
                    {
                        blankRowIndex = i;
                        break;
                    }
                }

                if (blankRowIndex != -1 && blankRowIndex + 1 < flowSheet.FlowRows.Count)
                {
                    for (int i = flowSheet.FlowRows.Count - 1; i >= blankRowIndex; i--)
                    {
                        flowSheet.FlowRows.RemoveAt(i);
                    }
                }
            }

            foreach (var bintableSheet in program.BintableSheets)
            {
                int blankRowIndex = -1;
                for (int i = 0; i < bintableSheet.BinTableRows.Count; i++)
                {
                    if (string.IsNullOrEmpty(bintableSheet.BinTableRows[i].Name) && string.IsNullOrEmpty(bintableSheet.BinTableRows[i].ItemList))
                    {
                        blankRowIndex = i;
                        break;
                    }
                }

                if (blankRowIndex != -1 && blankRowIndex + 1 < bintableSheet.BinTableRows.Count)
                {
                    for (int i = bintableSheet.BinTableRows.Count - 1; i >= blankRowIndex; i--)
                    {
                        bintableSheet.BinTableRows.RemoveAt(i);
                    }
                }
            }
        }
        // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item add end

        // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add start
        private List<string> GetItems(string line)
        {
            List<string> items = new List<string>();
            string item = "";
            for (int i = 0; i < line.Length; i++)
            {
                if (line.Substring(i, 1) == " " || line.Substring(i, 1) == "\t")
                {
                    items.Add(item);
                    item = "";
                }
                else
                {
                    item += line.Substring(i, 1);
                }
            }

            if (item != "")
            {
                items.Add(item);
            }
            return items;
        }
        // 2021-12-05  Bruce          #247           Fail_Stop_Table.txt support muti column and blanks on the bottom add end

        public void WorkFlow()
        {
            LoadConfig();

            //var exportfolder = Path.Combine(Directory.GetCurrentDirectory(), "tmp", Path.GetFileNameWithoutExtension(_tpFilePath) + "_ASCIIFiles_" + DateTime.Now.ToString("yyyyMMddHHmmssffffff"));
            var exportfolder = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "Teradyne", "PMICToolBox", "ExportTmp", "exportProg");
            if (!Directory.Exists(exportfolder))
                Directory.CreateDirectory(exportfolder);
            else
            {
                Directory.Delete(exportfolder, true);
                Directory.CreateDirectory(exportfolder);
            }
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
            //TestProgramUtility.ExportWorkBookCmd(_tpFilePath, exportfolder);
            //_testProgram = new IgxlProgram(_tpFilePath);
            TestProgramUtility.ExportWorkBookCmd(_guiArgs.TestPlanPath, exportfolder);
            _testProgram = new IgxlProgram(_guiArgs.TestPlanPath);
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
            _testProgram.LoadIgxlProgramAsync(exportfolder);

            // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item add start
            PreProcessData(_testProgram);
            // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item add end

            CreateResultFrame();
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
            //var ep = new ExcelPackage(new FileInfo(_outputFilePath));
            var ep = new ExcelPackage(new FileInfo(_guiArgs.OutputFilePath));
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end

            //Check Point 1
            CheckFailFlag();
            //Check Point2
            CheckBintableInFlow();
            //Check Point3
            CheckSetErrorBin();
            //Check Point4
            CheckSetBinInVBT();
            // Check Fail_Stop_Table
            CheckFail_Stop_Table();
            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
            CheckNoBinOut();
            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
            if (!string.IsNullOrEmpty(_guiArgs.DataLogPath))
            {
                CheckDatalogCompare();
            }
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

            // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
            AddUnusedBinTableRow();
            // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end

            OutputResult(ep);
        }

        private void OutputResult(ExcelPackage ep)
        {
            ep.Workbook.Worksheets.Add("BinOut");
            ep.Workbook.Worksheets["BinOut"].Cells.LoadFromCollection(_results, true);
            SetTittleFormat(ep.Workbook.Worksheets["BinOut"]);
            SetBinOutSheetErrorFormat(ep.Workbook.Worksheets["BinOut"]);

            ep.Workbook.Worksheets.Add("SetErrorBinReport");
            ep.Workbook.Worksheets["SetErrorBinReport"].Cells.LoadFromCollection(_setErrorBinResults, true);
            SetTittleFormat(ep.Workbook.Worksheets["SetErrorBinReport"]);
            SetSetErrorBinReportSheetErrorFormat(ep.Workbook.Worksheets["SetErrorBinReport"]);

            ep.Workbook.Worksheets.Add("FunctionBinReport");
            ep.Workbook.Worksheets["FunctionBinReport"].Cells.LoadFromCollection(_functionBinResults, true);
            SetTittleFormat(ep.Workbook.Worksheets["FunctionBinReport"]);

            ep.Workbook.Worksheets.Add("ErrorReport");
            ep.Workbook.Worksheets["ErrorReport"].Cells.LoadFromCollection(_errorResults, true);
            SetTittleFormat(ep.Workbook.Worksheets["ErrorReport"]);

            ep.Workbook.Worksheets.Add("VBTCheck");
            ep.Workbook.Worksheets["VBTCheck"].Cells.LoadFromCollection(_VBTSetBinCheckResults, true);
            SetTittleFormat(ep.Workbook.Worksheets["VBTCheck"]);

            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
            ep.Workbook.Worksheets.Add("NoBinOut");
            ep.Workbook.Worksheets["NoBinOut"].Cells.LoadFromCollection(_NoBinOutResultResults, true);
            SetTittleFormat(ep.Workbook.Worksheets["NoBinOut"]);
            // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
            if (!string.IsNullOrEmpty(_guiArgs.DataLogPath))
            {
                ep.Workbook.Worksheets.Add("NoBinOut ReferenceDatalog");
                ep.Workbook.Worksheets["NoBinOut ReferenceDatalog"].Cells.LoadFromCollection(_NoBinOutReferenceDatalogResults, true);
                SetTittleFormat(ep.Workbook.Worksheets["NoBinOut ReferenceDatalog"]);
            }
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

            foreach (var workSheet in ep.Workbook.Worksheets)
            {
                workSheet.Cells.AutoFitColumns();
            }

            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
            if (!string.IsNullOrEmpty(_guiArgs.DataLogPath))
            {
                SetDatalogCompareFormat(ep.Workbook.Worksheets["NoBinOut"], ep.Workbook.Worksheets["NoBinOut ReferenceDatalog"]);
            }
            // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start

            ep.Save();
            ep.Dispose();
        }

        private void SetTittleFormat(ExcelWorksheet worksheet)
        {
            for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                worksheet.Cells[1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.DarkOliveGreen);
                worksheet.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                worksheet.Cells[1, col].Style.Font.Size = 12;
                worksheet.Cells[1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
        }

        private void CheckFailFlag()
        {
            var errorResult = new ErrorResult();

            foreach (var flowSheet in _testProgram.FlowSheets)
            {
                foreach (var row in flowSheet.FlowRows)
                {
                    // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item del start
                    //// 2021-12-15  Bruce          #265	          Commend Flow add start
                    //if (string.IsNullOrEmpty(row.Opcode) && string.IsNullOrEmpty(row.Parameter))
                    //{
                    //    break;
                    //}
                    //// 2021-12-15  Bruce          #265	          Commend Flow add end
                    // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item del end

                    //Check Test Instance sort bin, hard bin and result have set
                    if (row.Opcode.Equals(Test))
                    {
                        if (!string.IsNullOrEmpty(row.BinFail) || !string.IsNullOrEmpty(row.SortFail) || !string.IsNullOrEmpty(row.Result))
                        {
                            _functionBinResults.Add(new FlowSetBinOutReport { FlowName = flowSheet.Name, Instance_Name = row.Parameter, SortBin = row.SortFail, HardBin = row.BinFail, Result = row.Result });
                        }
                    }

                    if (!row.Opcode.Equals(BinTable, StringComparison.OrdinalIgnoreCase)) continue;
                    var tmpBin = _results.FirstOrDefault(k => k.Name.Equals(row.Parameter, StringComparison.OrdinalIgnoreCase));
                    // Check Fail Flag content
                    if (tmpBin != null)
                    {
                        //tmpBin.FlowName.Add(flowSheet.Name);
                        var testName = row.Parameter.Replace("Bin_", "");
                        var instItem = flowSheet.FlowRows.FirstOrDefault(k => k.Opcode == Test && k.Parameter.Equals(testName));
                        // 2021-12-05  Bruce          #255	          Check Fail flag and BinTable Sequence Check by fail flag chg start
                        if (instItem == null)
                            instItem = flowSheet.FlowRows.FirstOrDefault(k => k.Opcode == Test && k.FailAction.Equals(tmpBin.ItemList));
                        // 2021-12-05  Bruce          #255	          Check Fail flag and BinTable Sequence Check by fail flag chg end
                        // Check Seq
                        if (instItem != null)
                        {
                            if (Convert.ToInt32(instItem.LineNum) > Convert.ToInt32(row.LineNum))
                                tmpBin.Fail_flag_and_BinTable_Sequence_Check = "Fail";
                            var binFlag = tmpBin.ItemList.Split(',').Select(p => p.Trim().ToLower()).ToList();
                            var flowFlag = instItem.FailAction.Split(',').Select(p => p.Trim().ToLower()).ToList();
                            var redundant = binFlag.Except(flowFlag);
                            var redundant2 = flowFlag.Except(binFlag);
                            if (redundant.Any() || redundant2.Any())
                            {
                                tmpBin.Fail_flag_and_Bin_define_in_same_flow = "Fail";
                            }
                        }
                        else
                        {
                            errorResult.SheetName = flowSheet.Name;
                            errorResult.Check = string.Format("Can not find Test Instance which use {0}", row.Parameter);
                        }
                    }
                    else
                    {
                        errorResult.SheetName = flowSheet.Name;
                        errorResult.Check = string.Format("Can not find {0} in BinTable", row.Parameter);
                    }
                }
            }
        }

        private void CheckBintableInFlow()
        {
            // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item chg start
            //// 2021-12-15  Bruce          #265	          Commend Flow chg start
            ////var allFlowRows = _testProgram.FlowSheets.SelectMany(p => p.FlowRows).ToList();
            //List<FlowRow> allFlowRows = new List<FlowRow>();
            //foreach (var flowsheet in _testProgram.FlowSheets)
            //{
            //    foreach (var row in flowsheet.FlowRows)
            //    {
            //        if (string.IsNullOrEmpty(row.Opcode) && string.IsNullOrEmpty(row.Parameter))
            //        {
            //            break;
            //        }
            //        allFlowRows.Add(row);
            //    }
            //}
            //// 2021-12-15  Bruce          #265	          Commend Flow chg end
            var allFlowRows = _testProgram.FlowSheets.SelectMany(p => p.FlowRows).ToList();
            // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item chg end

            foreach (var row in _results)
            {
                //list the flow name which has called this bin name
                var usedInflow = allFlowRows.FindAll(p => p.Opcode == BinTable && p.Parameter.Equals(row.Name, StringComparison.OrdinalIgnoreCase)).ToList();
                var flowNames = "";
                if (usedInflow.Any())
                {
                    var tmp = usedInflow.Select(flowRow => flowRow.SheetName).ToList();
                    var removeRepeat = tmp.Distinct();
                    flowNames = string.Join(",", removeRepeat);
                }
                row.FlowName = flowNames;

                //Check Sort Bin duplicate
                var repeat = _results.FindAll(p => p.Sort.Equals(row.Sort)).ToList();
                if (repeat.Count > 1)
                {
                    row.SwBinDuplicate = "Y";
                }

                //BinOut in each subFlow
                var targetFlow = allFlowRows.FirstOrDefault(p => p.Opcode == BinTable && p.Parameter.Equals(row.Name, StringComparison.OrdinalIgnoreCase));
                if (targetFlow == null)
                {
                    row.Redundant_Bin = "Not Exist in Any SubFlow";
                }

                //Check Fail Flag in Same Flow
                if (targetFlow == null || targetFlow.SheetName.Contains("Init_EnableWd")) continue;
                var flags = row.ItemList.Split(',').ToList();

                var usedFlag = false;
                var checkTestItem = allFlowRows.FindAll(p => p.SheetName.Equals(targetFlow.SheetName));
                flags.ForEach(k =>
                {
                    if (checkTestItem.Exists(p => p.FailAction.ToLower().Contains(k.ToLower())))
                    {
                        usedFlag = true;
                    }
                });

                // 2021-12-08  Bruce          #256	          Check Flag Default Value chg start
                //if (usedFlag) continue;
                //row.Fail_flag_and_Bin_define_in_same_flow += ",Fail flag is not used in SubFlow";
                //row.Fail_flag_and_Bin_define_in_same_flow = row.Fail_flag_and_Bin_define_in_same_flow.TrimStart(',');
                if (!usedFlag)
                {
                    row.Fail_flag_and_Bin_define_in_same_flow += ",Fail flag is not used in SubFlow";
                    row.Fail_flag_and_Bin_define_in_same_flow = row.Fail_flag_and_Bin_define_in_same_flow.TrimStart(',');
                }
                // 2021-12-08  Bruce          #256	          Check Flag Default Value chg end

                // 2021-12-08  Bruce          #256	          Check Flag Default Value add start
                BinTableRow binRow = GetBinTableRow(row.Name);
                if (binRow != null)
                {
                    // 2021-12-15  Bruce          #263	          Error Message Modify: Fail Flag count is not match with item‘s count ->FlagDefaultValue Error chg start
                    //if (binRow.Result.Equals("Fail", StringComparison.InvariantCultureIgnoreCase))
                    //{
                    //    if (!ContainsPossibleValues(binRow.Items, new List<string>() { "F", "T" }))
                    //    {
                    //        if (!string.IsNullOrEmpty(row.Fail_flag_and_Bin_define_in_same_flow)) row.Fail_flag_and_Bin_define_in_same_flow += ",";
                    //        row.Fail_flag_and_Bin_define_in_same_flow += "only F or T when result is Fail";
                    //    }
                    //}
                    //else if (binRow.Result.Equals("Fail-stop", StringComparison.InvariantCultureIgnoreCase))
                    //{
                    //    if (!ContainsPossibleValues(binRow.Items, new List<string>() { "T" }))
                    //    {
                    //        if (!string.IsNullOrEmpty(row.Fail_flag_and_Bin_define_in_same_flow)) row.Fail_flag_and_Bin_define_in_same_flow += ",";
                    //        row.Fail_flag_and_Bin_define_in_same_flow += "only T when result is Fail-stop";
                    //    }
                    //}

                    //int itemsCount = binRow.ItemList.Split(',').Length;
                    //if (itemsCount != binRow.Items.Count)
                    //{
                    //    if (!string.IsNullOrEmpty(row.Fail_flag_and_Bin_define_in_same_flow)) row.Fail_flag_and_Bin_define_in_same_flow += ",";
                    //    row.Fail_flag_and_Bin_define_in_same_flow += "Fail Flag count is not match with item's count";
                    //}

                    if (!string.IsNullOrEmpty(binRow.ItemList) && binRow.Items.Count == 0)
                    {
                        row.Symbol_Define_Of_Fail_Flag = "Items is empty";
                        continue;
                    }

                    int itemsCount = binRow.ItemList.Split(',').Length;
                    if (itemsCount != binRow.Items.Count)
                    {
                        row.Symbol_Define_Of_Fail_Flag = "FlagDefaultValue Error";
                        continue;
                    }

                    if (binRow.Result.Equals("Fail", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!ContainsPossibleValues(binRow.Items, new List<string>() { "F", "T" }))
                        {
                            row.Symbol_Define_Of_Fail_Flag = "only F or T when result is Fail";
                        }
                    }
                    else if (binRow.Result.Equals("Fail-stop", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!ContainsPossibleValues(binRow.Items, new List<string>() { "T" }))
                        {
                            row.Symbol_Define_Of_Fail_Flag = "only T when result is Fail-stop";
                        }
                    }
                    // 2021-12-15  Bruce          #263	          Error Message Modify: Fail Flag count is not match with item‘s count ->FlagDefaultValue Error chg end
                }
                // 2021-12-08  Bruce          #256	          Check Flag Default Value add end
            }
        }

        private void CheckSetErrorBin()
        {
            foreach (var flowSheet in _testProgram.FlowSheets)
            {
                var setErrorBin = new SetErrorBinResult { FlowName = flowSheet.Name };
                // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item chg start
                //// 2021-12-15  Bruce          #265	          Commend Flow chg start
                ////var setErrorBinItem = flowSheet.FlowRows.FindAll(p => p.Opcode.Equals(SetErrorBin, StringComparison.OrdinalIgnoreCase));
                //List<FlowRow> setErrorBinItem = new List<FlowRow>();
                //foreach (var row in flowSheet.FlowRows)
                //{
                //    if (string.IsNullOrEmpty(row.Opcode) && string.IsNullOrEmpty(row.Parameter))
                //    {
                //        break;
                //    }
                //    if (row.Opcode.Equals(SetErrorBin, StringComparison.OrdinalIgnoreCase))
                //    {
                //        setErrorBinItem.Add(row);
                //    }
                //}
                //// 2021-12-15  Bruce          #265	          Commend Flow chg end
                var setErrorBinItem = flowSheet.FlowRows.FindAll(p => p.Opcode.Equals(SetErrorBin, StringComparison.OrdinalIgnoreCase));
                // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item chg end

                if (setErrorBinItem.Any())
                {
                    if (setErrorBinItem.Count > 1)
                    {
                        setErrorBin.Is_Error_Bin_number_Define_in_Bin_table = "redundant";
                        continue;
                    }
                    var target = _results.FirstOrDefault(k => k.Sort.Equals(setErrorBinItem[0].SortFail) && k.Bin.Equals(setErrorBinItem[0].BinFail));
                    if (target != null)
                    {
                        setErrorBin.Is_Error_Bin_number_Define_in_Bin_table = "Pass";
                        setErrorBin.Error_Sort_Bin_Number = setErrorBinItem[0].SortFail;
                        setErrorBin.Error_Hard_Bin_Number = setErrorBinItem[0].BinFail;
                        setErrorBin.Bin_Name = target.Name;
                    }
                    else
                    {
                        setErrorBin.Is_Error_Bin_number_Define_in_Bin_table = "Fail";
                        setErrorBin.Error_Sort_Bin_Number = setErrorBinItem[0].SortFail;
                        setErrorBin.Error_Hard_Bin_Number = setErrorBinItem[0].BinFail;
                    }
                }
                else
                    setErrorBin.Is_Error_Bin_number_Define_in_Bin_table = "Not exist set-error-bin";

                _setErrorBinResults.Add(setErrorBin);
            }
            _setErrorBinResults.Sort((x, y) => Convert.ToInt32(x.Error_Sort_Bin_Number).CompareTo(Convert.ToInt32(y.Error_Sort_Bin_Number)));
        }

        private void CheckSetBinInVBT()
        {
            foreach (var module in _testProgram.Modules)
            {
                _VBTSetBinCheckResults.AddRange(CheckVBTFile(module));
            }
        }

        private void CheckFail_Stop_Table()
        {
            foreach (var binOutCheckResult in _results)
            {
                BinOutConfig binoutcfg = _binoutCfg.FirstOrDefault(b => b.Sort == binOutCheckResult.Sort);
                if (binoutcfg == null)
                {
                    continue;
                }

                string errorMsg = "";
                // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
                binoutcfg.IsUsed = true;
                // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
                if (!binOutCheckResult.Name.Equals(binoutcfg.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    errorMsg = string.Format("Bin Name Mismatch(correct : {0})", binoutcfg.Name);
                }
                if (!binOutCheckResult.ItemList.Equals(binoutcfg.ItemList, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (errorMsg != "") errorMsg += ",";
                    errorMsg += string.Format("Fail Flag Mismatch(correct : {0})", binoutcfg.ItemList);
                }
                if (!binOutCheckResult.Result.Equals(binoutcfg.Result))
                {
                    if (errorMsg != "") errorMsg += ",";
                    errorMsg += string.Format("Result Mismatch(correct : {0})", binoutcfg.Result);
                }

                if (errorMsg != "")
                {
                    binOutCheckResult.Standard_BinTable_Mismatch = errorMsg;
                }
            }
        }

        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
        private void CheckNoBinOut()
        {
            foreach (var flowSheet in _testProgram.FlowSheets)
            {
                foreach (var row in flowSheet.FlowRows)
                {
                    // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item del start
                    //// 2021-12-15  Bruce          #265	          Commend Flow add start
                    //if (string.IsNullOrEmpty(row.Opcode) && string.IsNullOrEmpty(row.Parameter))
                    //{
                    //    break;
                    //}
                    //// 2021-12-15  Bruce          #265	          Commend Flow add end
                    // 2021-12-17  Bruce          #267	          Commend In BinTable Sheet：Remove Commend Item del end

                    //Check Test Instance sort bin, hard bin and result have set
                    if (row.Opcode.Equals(Test))
                    {
                        string Comment = "";

                        if (row.FailAction == "")
                        {
                            // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg start
                            //Comment = "NoBinOut";
                            Comment = "Without Pass/Fail Flag Define";
                            // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg end
                        }
                        else
                        {
                            // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg start
                            // BinTableRow binRow = GetBinTableRowByFailFlag(row.FailAction);
                            // if (binRow == null)
                            List<BinTableRow> binRows = GetBinTableRowByFailFlag(row.FailAction);
                            if (!binRows.Any())
                            {
                            // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg end

                                // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg start
                                //Comment = "FlagWithoutBinTable";
                                Comment = "Can't Find Any BinTable Used this Flag";
                                // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg end
                            }
                            else
                            {
                                // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg start
                                //FlowRow binoutFlowRow = flowSheet.FlowRows.FirstOrDefault(r => r.Parameter.Equals(binRow.Name, StringComparison.InvariantCultureIgnoreCase));
                                //if (binoutFlowRow == null)
                                //{
                                //    // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg start
                                //    //Comment = "NoBinOutInFlow";
                                //    Comment = "Without Match BinTable In Subflow";
                                //    // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg end
                                //}
                                //else
                                //{
                                //    if (binoutFlowRow.Enable != "" ||
                                //        binoutFlowRow.Job != "" ||
                                //        binoutFlowRow.Part != "" ||
                                //        binoutFlowRow.Env != "")
                                //    {
                                //        // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg start
                                //        //Comment = "BinTableWithCondition";
                                //        Comment = string.Format("\"{0}\" With Gate Condition", binRow.Name);
                                //        // 2021-12-15  Bruce          #264	          No Bin Out Error Message Modify chg end
                                //    }
                                //}

                                List<string> notMatchBinTableList = new List<string>();
                                List<string> withGateConditionBinTableList = new List<string>();

                                binRows.ForEach(delegate (BinTableRow binRow)
                                {
                                    FlowRow binoutFlowRow = flowSheet.FlowRows.FirstOrDefault(r => r.Parameter.Equals(binRow.Name, StringComparison.InvariantCultureIgnoreCase));
                                    if (binoutFlowRow == null)
                                    {
                                        notMatchBinTableList.Add(binRow.Name);
                                    }
                                    else
                                    {
                                        if (binoutFlowRow.Enable != "" ||
                                            binoutFlowRow.Job != "" ||
                                            binoutFlowRow.Part != "" ||
                                            binoutFlowRow.Env != "")
                                        {
                                            withGateConditionBinTableList.Add(binRow.Name);
                                        }
                                    }
                                });
                                if (notMatchBinTableList.Any())
                                    Comment = string.Format("\"{0}\" Without Match BinTable In Subflow. ", string.Join(",", notMatchBinTableList));
                                if (withGateConditionBinTableList.Any())
                                    Comment += string.Format("\"{0}\" With Gate Condition.", string.Join(",", withGateConditionBinTableList));
                                // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg end
                            }
                        }

                        if (Comment != "")
                            _NoBinOutResultResults.Add(new NoBinOutResult() { FlowSheet = flowSheet.Name, TestItem = row.Parameter, Comment = Comment });
                    }
                }
            }
        }
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
        private void CheckDatalogCompare()
        {
            // 2022-02-21  Steven Chen    #320            Support No FlowHeader in Binoutcheck add start
            //foreach (var flow in _DataLogFileInfo.TestFlows.Keys)
            //{
            //    List<string> testItems = _DataLogFileInfo.TestFlows[flow];
            //    if (testItems.Count > 0)
            //    {
            //        foreach (var item in testItems)
            //        {
            //            NoBinOutResult res = _NoBinOutResultResults.FirstOrDefault(R => R.FlowSheet.Equals("Flow_" + flow, StringComparison.InvariantCultureIgnoreCase)
            //                                                && R.TestItem.Equals(item, StringComparison.InvariantCultureIgnoreCase));
            //            if (res != null)
            //            {
            //                _NoBinOutReferenceDatalogResults.Add(new NoBinOutReferenceDatalogResult()
            //                {
            //                    FlowSheet = res.FlowSheet,
            //                    TestItem = res.TestItem,
            //                    Comment = res.Comment
            //                });
            //            }
            //        }
            //    }
            //}
            foreach (var flowStep in _DataLogFileInfo.FlowSteps)
            {
                //test flow table
                if (flowStep.Value)
                {
                    List<string> testItems = _DataLogFileInfo.TestFlows[flowStep.Key];
                    if (testItems.Count > 0)
                    {
                        foreach (var item in testItems)
                        {
                            NoBinOutResult res = _NoBinOutResultResults.FirstOrDefault(R => R.FlowSheet.Equals("Flow_" + flowStep.Key, StringComparison.InvariantCultureIgnoreCase)
                                                                && R.TestItem.Equals(item, StringComparison.InvariantCultureIgnoreCase));
                            if (res != null)
                            {
                                _NoBinOutReferenceDatalogResults.Add(new NoBinOutReferenceDatalogResult()
                                {
                                    FlowSheet = res.FlowSheet,
                                    TestItem = res.TestItem,
                                    Comment = res.Comment
                                });
                            }
                        }
                    }
                }
                //test item
                else
                {
                    var results = _NoBinOutResultResults.FindAll(result => result.TestItem.Equals(flowStep.Key, StringComparison.InvariantCultureIgnoreCase));
                    NoBinOutReferenceDatalogResult datalogResult = new NoBinOutReferenceDatalogResult();
                    _NoBinOutReferenceDatalogResults.Add(datalogResult);
                    datalogResult.TestItem = flowStep.Key;
                    datalogResult.ErrorComment = "Datalog without \"Flow_Start\"";

                    if (results.Count == 0)
                    {
                        datalogResult.FlowSheet = string.Empty;
                        datalogResult.Comment = string.Empty;
                    }
                    else if (results.Count == 1)
                    {
                        datalogResult.FlowSheet = results.FirstOrDefault().FlowSheet;
                        datalogResult.Comment = results.FirstOrDefault().Comment;
                    }
                    else if (results.Count > 1)
                    {
                        datalogResult.FlowSheet = string.Join(";", results.Select(o => o.FlowSheet));
                        datalogResult.Comment = string.Join(";", results.Select(o => o.Comment));
                    }
                }
            }
            // 2022-02-21  Steven Chen    #320            Support No FlowHeader in Binoutcheck add end

            _NoBinOutReferenceDatalogResults.Add(new NoBinOutReferenceDatalogResult());

            _NoBinOutReferenceDatalogResults.Add(new NoBinOutReferenceDatalogResult()
            {
                FlowSheet = _guiArgs.DataLogPath
            });
        }
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

        Regex RegFunStart = new Regex(@"((Function)|(Sub)) (?<name>\w+)(\()", RegexOptions.IgnoreCase);
        Regex RegFunEnd = new Regex(@"(End )((Function)|(Sub))", RegexOptions.IgnoreCase);
        Regex RegSortBinNumber = new Regex(@"(\.)(SortNumber =)", RegexOptions.IgnoreCase);
        Regex RegHardBinNumber = new Regex(@"(\.)(BinNumber =)", RegexOptions.IgnoreCase);

        private List<VBTSetBinCheckResult> CheckVBTFile(string vbtFile)
        {
            List<VBTSetBinCheckResult> resultList = new List<VBTSetBinCheckResult>();
            string[] allLines = File.ReadAllLines(vbtFile);
            FileInfo fi = new FileInfo(vbtFile);
            string moduleName = fi.Name;
            string methodName = "";
            bool sortBin = false;
            bool hardBin = false;
            for (int i = 0; i < allLines.Length; i++)
            {
                string line = allLines[i];
                Match lresult = RegFunStart.Match(line);
                if (lresult.Success)
                {
                    methodName = lresult.Groups["name"].ToString();
                    sortBin = false;
                    hardBin = false;
                    continue;
                }
                lresult = RegFunEnd.Match(line);
                if (lresult.Success)
                {
                    if (methodName != "" && (sortBin == true || hardBin == true))
                    {
                        resultList.Add(new VBTSetBinCheckResult() { Function = methodName, Module_Name = moduleName, Hard = hardBin, Sort = sortBin });
                    }
                    continue;
                }

                if (RegSortBinNumber.IsMatch(line))
                {
                    sortBin = true;
                    continue;
                }

                if (RegHardBinNumber.IsMatch(line))
                {
                    hardBin = true;
                    continue;
                }
            }

            return resultList;
        }

        private void SetBinOutSheetErrorFormat(ExcelWorksheet worksheet)
        {
            bool hasError = false;

            // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
            bool standardBinMode = false;
            // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end

            // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add start
            int CheckStartColumn = 8;
            int CheckEndColumn = worksheet.Dimension.End.Column;
            // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add end
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
                if (!standardBinMode)
                {
                    bool isblankRow = true;
                    for (int r = 1; r < 8; r++)
                    {
                        if (!string.IsNullOrEmpty(worksheet.Cells[i, r].Text))
                        {
                            isblankRow = false;
                            break;
                        }
                    }
                    if (isblankRow)
                    {
                        standardBinMode = true;
                        continue;
                    }

                    // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
                    hasError = false;
                    string flowName = worksheet.Cells[i, 7].Text;
                    if (flowName.IndexOf(",") >= 0)
                    {
                        hasError = true;
                    }

                    // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg start
                    //// 2021-12-05  Bruce          #254	          Missing Color in column SwBinDuplicate chg start
                    ////if (worksheet.Cells[i, 8].Text != "" ||
                    ////    worksheet.Cells[i, 9].Text != "" ||
                    ////    worksheet.Cells[i, 10].Text != "" ||
                    ////    worksheet.Cells[i, 12].Text != "")
                    //if (worksheet.Cells[i, 8].Text != "" ||
                    //    worksheet.Cells[i, 9].Text != "" ||
                    //    worksheet.Cells[i, 10].Text != "" ||
                    //    worksheet.Cells[i, 11].Text != "" ||
                    //    worksheet.Cells[i, 12].Text != "")
                    //// 2021-12-05  Bruce          #254	          Missing Color in column SwBinDuplicate chg end
                    //{
                    //    hasError = true;
                    //}

                    for (int c = CheckStartColumn; c <= CheckEndColumn; c++)
                    {
                        if (worksheet.Cells[i, c].Text != "")
                        {
                            hasError = true;
                            break;
                        }
                    }
                    // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg end

                    if (worksheet.Cells[i, 4].Text.Equals("Fail", StringComparison.InvariantCultureIgnoreCase) ||
                        worksheet.Cells[i, 4].Text.Equals("Not exist set-error-bin", StringComparison.InvariantCultureIgnoreCase))
                    {
                        hasError = true;
                    }

                    if (hasError)
                    {
                        // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg start
                        //worksheet.Row(i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //worksheet.Row(i).Style.Fill.BackgroundColor.SetColor(Color.Red);
                        ExcelRange errorRow = worksheet.Cells[i, 1, i, CheckEndColumn];
                        errorRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        errorRow.Style.Fill.BackgroundColor.SetColor(ErrorColor);
                        // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg end
                    }
                    // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
                }
                else
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[i, 1].Text))
                    {
                        ExcelRange errorRow = worksheet.Cells[i, 1, i, CheckEndColumn];
                        errorRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        errorRow.Style.Fill.BackgroundColor.SetColor(ErrorColor2);
                    }
                }
                // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
            }
        }

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
        private void SetDatalogCompareFormat(ExcelWorksheet worksheetbase, ExcelWorksheet worksheetDatalog)
        {
            worksheetDatalog.Column(1).Width = worksheetbase.Column(1).Width;
        }
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end

        private void SetSetErrorBinReportSheetErrorFormat(ExcelWorksheet worksheet)
        {
            bool hasError = false;
            // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add start
            int CheckEndColumn = worksheet.Dimension.End.Column;
            // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area add end
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                hasError = false;
                if (worksheet.Cells[i, 4].Text.Equals("Fail", StringComparison.InvariantCultureIgnoreCase) ||
                    worksheet.Cells[i, 4].Text.Equals("Not exist set-error-bin", StringComparison.InvariantCultureIgnoreCase))
                {
                    hasError = true;
                }

                if (hasError)
                {
                    // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg start
                    //worksheet.Row(i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //worksheet.Row(i).Style.Fill.BackgroundColor.SetColor(Color.Red);
                    ExcelRange errorRow = worksheet.Cells[i, 1, i, CheckEndColumn];
                    errorRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    errorRow.Style.Fill.BackgroundColor.SetColor(ErrorColor);
                    // 2021-12-15  Bruce          #262	          Set New Color and adjust Color area chg end
                }
            }
        }

        // 2021-12-08  Bruce          #256	          Check Flag Default Value add start
        private BinTableRow GetBinTableRow(string name)
        {
            BinTableRow result = null;
            foreach (var sheet in _testProgram.BintableSheets)
            {
                result = sheet.BinTableRows.FirstOrDefault(r => r.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        private bool ContainsPossibleValues(List<string> items, List<string> possilbeValues)
        {
            foreach (var item in items)
            {
                if (!possilbeValues.Contains(item, _StringIgnoreCaseCompare)) return false;
            }
            return true;
        }
        // 2021-12-08  Bruce          #256	          Check Flag Default Value add end


        // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg start
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add start
        //private BinTableRow GetBinTableRowByFailFlag(string failflag)
        //{
        //    BinTableRow result = null;
        //    foreach (var sheet in _testProgram.BintableSheets)
        //    {
        //        result = sheet.BinTableRows.FirstOrDefault(r => r.ItemList.Split(',').Contains(failflag, _StringIgnoreCaseCompare));
        //        if (result != null)
        //        {
        //            return result;
        //        }
        //    }

        //    return null;
        //}
        // 2021-12-08  Bruce          #258	          Add No Bin Out Check add end

        private List<BinTableRow> GetBinTableRowByFailFlag(string failflag)
        {
            List<BinTableRow> result = new List<BinTableRow>();
            foreach (var sheet in _testProgram.BintableSheets)
            {
                result.AddRange(sheet.BinTableRows.FindAll(r => r.ItemList.Split(',').Contains(failflag, _StringIgnoreCaseCompare)));
            }
            return result.Distinct().ToList();
        }
        // 2022-01-24  Steven Chen    #304            No bin out check all action-fail flag related bintables. chg end 

        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
        private void AddUnusedBinTableRow()
        {
            if (_binoutCfg.Exists(b => b.IsUsed == false))
            {
                _results.Add(new BinOutCheckResult());

                for (int i = 0; i < _binoutCfg.Count; i++)
                {
                    if (_binoutCfg[i].IsUsed == false)
                    {
                        var binOutCheckResult = new BinOutCheckResult();
                        binOutCheckResult.Name = _binoutCfg[i].Name;
                        binOutCheckResult.ItemList = _binoutCfg[i].ItemList;
                        binOutCheckResult.Op = _binoutCfg[i].Op;
                        binOutCheckResult.Sort = _binoutCfg[i].Sort;
                        binOutCheckResult.Bin = _binoutCfg[i].Bin;
                        binOutCheckResult.Result = _binoutCfg[i].Result;
                        binOutCheckResult.Standard_BinTable_Mismatch = "Standard Bin Not Exist in BinTable";
                        _results.Add(binOutCheckResult);
                    }
                }
            }
        }
        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
    }
}
