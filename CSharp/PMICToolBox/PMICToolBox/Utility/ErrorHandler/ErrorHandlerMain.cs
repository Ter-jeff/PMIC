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
// 2022-02-11  Steven Chen    #310            Simplify Output Option ErrorHandler sheet(Remove Pass Item)
// 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item)
//------------------------------------------------------------------------------ 
using PmicAutomation.MyControls;
using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace PmicAutomation.Utility.ErrorHandler
{
    public class ErrorHandlerMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText = null;
        //private readonly List<string> _inputFile;
        private List<Result> _result;
        private List<OptionExplicitResult> _oeResult;
        private readonly string _inputPath;
        private readonly string _outputFolder;
        private string _outputFilePath;
        private bool GenNewBas;
        private const string OnErrorResumeNext = "On Error Resume Next";
        private const string OnErrorGoTo1 = "On Error GoTo errHandler";
        private const string OnErrorGoTo2 = "    Dim sCurrentFuncName As String:: sCurrentFuncName = \"";
        private const string ErrorHandler1 = "errHandler:";
        private const string ErrorHandler2 = "    TheExec.AddOutput \"<Error> \" + sCurrentFuncName + \":: please Check it out.\"";
        private const string ErrorHandler3 = "    TheExec.Datalog.WriteComment \"<Error> \" + sCurrentFuncName + \":: please check it out.\"";
        private const string ErrorHandler4 = "    If AbortTest Then Exit ";

        public ErrorHandlerMain(ErrorHandlerForm errorHandler)
        {
            _appendText = errorHandler.AppendText;
            _inputPath = errorHandler.FileOpen_ErrorHandler.ButtonTextBox.Text;
            _outputFolder = errorHandler.FileOpen_OutputPath.ButtonTextBox.Text;
            _outputFilePath = Path.Combine(_outputFolder, "Report_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            _result = new List<Result>();
            _oeResult = new List<OptionExplicitResult>();
            GenNewBas = errorHandler.chkAddErrHandler.Checked;
        }

        public ErrorHandlerMain(string inputPath, string outputFilePath, bool chkAddErrHandler)
        {
            _inputPath = inputPath;

            if (!Directory.Exists(new FileInfo(outputFilePath).DirectoryName))
                throw new Exception("Output file parent folder is not exist: " + new FileInfo(outputFilePath).DirectoryName);
            if (!outputFilePath.EndsWith(".xlsx"))
                throw new Exception("Output file is not an xlsx excel file: " + outputFilePath);
            _outputFolder = new FileInfo(outputFilePath).DirectoryName;
            _outputFilePath = outputFilePath;
            GenNewBas = chkAddErrHandler;
            _result = new List<Result>();
            _oeResult = new List<OptionExplicitResult>();
        }

        public void WorkFlow()
        {
            List<string> files;

            if (Directory.Exists(_inputPath))
            {
                files = GetAllVbtModuleFilesFromDirectory(_inputPath);
            }
            else
            {
                files = GetAllVbtModuleFilesFromTpFile(_inputPath);
            }

            foreach (var file in files)
            {
                var fileName = Path.GetFileName(file);
                var dspProcedure = false;
                if (Regex.IsMatch(fileName.ToUpper(), "^DSP_"))
                    dspProcedure = true;

                WriteUILog("Process " + fileName, Color.Black);
                if (Path.GetExtension(file) == "igxl")
                {

                }
                else
                {
                    var content = File.ReadAllLines(file).ToList();

                    bool existOptionExplicit = content.Exists(p => p.Trim().StartsWith("Option Explicit", StringComparison.OrdinalIgnoreCase));
                    if (!existOptionExplicit)
                    {
                        int idx = content.FindIndex(p => p.Trim().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase));
                        content.Insert(idx + 1, "Option Explicit");
                    }
                    // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg start
                    //_oeResult.Add(new OptionExplicitResult { FileName = Path.GetFileName(file), ExistOptionExplicit = existOptionExplicit });
                    if (!existOptionExplicit)
                    {
                        _oeResult.Add(new OptionExplicitResult { FileName = Path.GetFileName(file) });
                    }
                    // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg end

                    var procedures = new List<Procedure>();
                    var originalProcedures = new List<Procedure>(); // used for dump report
                    content = FilterOutDummyLine(content, procedures, originalProcedures);
                    var ProcedureList = BasFileManage.GetProcedureStartEndPosition(content, procedures);
                    var missErrHandler = ProcedureList.FindAll(p => !p.ExistOnErrorGoTo).ToList();

                    bool bEverModifiedErrorName = false;
                    originalProcedures.ForEach(x =>
                    {
                        if (dspProcedure)
                        {
                            // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add start
                            if (!x.ExistOnErrorResumeNext || x.StopKeyword || x.MsgBoxKeyword)
                            // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add end
                                _result.Add(new Result
                                {
                                    FileName = Path.GetFileName(file),
                                    FuncName = x.Name,
                                    Modified = !x.ExistOnErrorResumeNext,
                                    //ExistErrHandler = x.ExistErrHandler
                                    StopKeyword = x.StopKeyword,
                                    MsgBoxKeyword = x.MsgBoxKeyword
                                });
                        }
                        else
                        {
                            // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add start
                            if ((!x.ExistOnErrorGoTo || x.ModifiedErrorName) || x.StopKeyword || x.MsgBoxKeyword)
                            // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add end
                                _result.Add(new Result
                                {
                                    FileName = Path.GetFileName(file),
                                    FuncName = x.Name,
                                    Modified = !x.ExistOnErrorGoTo || x.ModifiedErrorName,
                                    //ExistErrHandler = x.ExistErrHandler
                                    StopKeyword = x.StopKeyword,
                                    MsgBoxKeyword = x.MsgBoxKeyword
                                });
                        }
                        if (x.ModifiedErrorName)
                            bEverModifiedErrorName = true;
                    });

                    if (GenNewBas && (missErrHandler.Any() || bEverModifiedErrorName))
                    {
                        var output = Path.Combine(_outputFolder, Path.GetFileName(file));
                        if (!Directory.Exists(_outputFolder))
                            Directory.CreateDirectory(_outputFolder);
                        StreamWriter sw = new StreamWriter(output);
                        for (int i = 0; i < content.Count; i++)
                        {
                            if (missErrHandler.Any(x => x.End == i))
                            {
                                var target = missErrHandler.FirstOrDefault(p => p.End == i);
                                if (!dspProcedure)
                                {
                                    if (!target.ExistOnErrorGoTo && !target.ExistErrHandler)
                                    {
                                        sw.WriteLine("Exit " + target.Type);
                                        sw.WriteLine(ErrorHandler1);
                                        sw.WriteLine(ErrorHandler2);
                                        sw.WriteLine(ErrorHandler3);
                                        sw.WriteLine(ErrorHandler4 + target.Type + " Else Resume Next");
                                    }
                                }
                                else
                                {
                                    //if (!target.ExistOnErrorGoTo)
                                    //    sw.WriteLine("Exit " + target.Type);
                                }
                            }

                            sw.WriteLine(content[i]);

                            if (missErrHandler.Any(x => x.Start == i))
                            {
                                var target = missErrHandler.FirstOrDefault(x => x.Start == i);
                                if (!dspProcedure)
                                {
                                    if (!target.ExistOnErrorGoTo)
                                    {
                                        sw.WriteLine(OnErrorGoTo1);
                                        if (!target.ExistFuncName)
                                            sw.WriteLine(OnErrorGoTo2 + target.Name + "\"");
                                    }
                                }
                                else
                                {
                                    if (!target.ExistOnErrorResumeNext)
                                    {
                                        sw.WriteLine(OnErrorResumeNext);
                                    }
                                }
                            }
                        }
                        sw.Close();
                    }
                }
            }

            WriteReport();
            WriteUILog("All processes were completed !!!", Color.Black);
        }

        private List<string> GetAllVbtModuleFilesFromDirectory(string inputPath)
        {
            List<string> igxlFiles = Directory.GetFiles(_inputPath, "*").Where(p => p.EndsWith(".igxl", StringComparison.OrdinalIgnoreCase)).ToList();

            if (igxlFiles.Count > 0 && FindInstalledOasis() == false)
            {
                WriteUILog("NO Oasis installed !!! We can only work with .cls and .bas files.", Color.Green);
            }
            else
            {
                foreach (var file in igxlFiles)
                {
                    string exportASCIIFilesPath = Path.Combine(inputPath, "tmp", Path.GetFileNameWithoutExtension(file) + "_ASCIIFiles");
                    ExportWorkBook(file, exportASCIIFilesPath);
                }
            }

            List<string> files = Directory.GetFiles(inputPath, "*", SearchOption.AllDirectories).Where(p => p.EndsWith(".bas", StringComparison.OrdinalIgnoreCase)
            || p.EndsWith(".cls", StringComparison.OrdinalIgnoreCase)).ToList();

            return files;
        }

        private List<string> GetAllVbtModuleFilesFromTpFile(string inputPath)
        {
            string exportASCIIFilesPath;
            if (FindInstalledOasis() == false)
            {
                throw new Exception("NO Oasis installed !!!");
            }
            else
            {
                exportASCIIFilesPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "Teradyne", "PMICToolBox", "ExportTmp", "exportProg");
                ExportWorkBook(inputPath, exportASCIIFilesPath);
            }

            List<string> files = Directory.GetFiles(exportASCIIFilesPath, "*", SearchOption.AllDirectories).Where(p => p.EndsWith(".bas", StringComparison.OrdinalIgnoreCase)
            || p.EndsWith(".cls", StringComparison.OrdinalIgnoreCase)).ToList();

            return files;
        }

        private static bool FindInstalledIgxl()
        {
            string igxlRoot = Environment.GetEnvironmentVariable("IGXLROOT");
            if (string.IsNullOrEmpty(igxlRoot))
                return false;
            else
                return true;
        }

        private static bool FindInstalledOasis()
        {
            string oasisRoot = Environment.GetEnvironmentVariable("OASISROOT");
            if (string.IsNullOrEmpty(oasisRoot))
                return false;
            else
                return true;
        }

        private bool ExportWorkBook(string testProgramName, string exportFolder)
        {
            if (!Directory.Exists(exportFolder))
                Directory.CreateDirectory(exportFolder);
            else
            {
                Directory.Delete(exportFolder, true);
                Directory.CreateDirectory(exportFolder);
            }
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
            if (File.Exists(exportWorkBookCmd))
            {
                string option = "-w \"" + testProgramName + "\" -d \"" + exportFolder + "\"";
                return RunCmd("\"" + exportWorkBookCmd + "\"", option);
            }
            return false;
        }

        private bool RunCmd(string cmd, string argument = "")
        {
            Process nProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = cmd;
            startInfo.Arguments = argument;
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
            if (nProcess.ExitCode == 0)
                return true;
            return false;
        }

        private List<string> FilterOutDummyLine(List<string> lines, List<Procedure> procedures, List<Procedure> originalProcedures)
        {
            List<string> result = new List<string>();
            bool multiLine = false;

            for (int i = 0; i < lines.Count; ++i)
            {
                string line = BasFileManage.RemoveComment(lines[i]);
                string type;
                string subName = BasFileManage.GetFunctionName(line, out type);
                string onErrorLine = string.Empty;
                string exitFunctionLine = string.Empty;
                string label = "errHandler"; // default label name
                bool bOnErrorGoTo = false;
                bool bExitFunction = false;

                result.Add(lines[i]);

                if (string.IsNullOrEmpty(type) || type == "Enum" || subName.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (line.EndsWith(" _"))
                    multiLine = true;

                List<string> errHandlerList = new List<string>();

                List<string> tmpResult = new List<string>();
                var procedure = new Procedure { Name = subName, Type = type };
                int firstOnErrorGoToIndex = -1;
                for (int j = i + 1; j < lines.Count; j++)
                {
                    line = BasFileManage.RemoveComment(lines[j]);
                    tmpResult.Add(lines[j]);
                    if (multiLine)
                    {
                        if (Regex.IsMatch(line.TrimEnd(), @"\)") && !line.EndsWith(" _"))
                        {
                            procedure.Start = j;
                            multiLine = false;
                        }
                    }

                    if (string.IsNullOrEmpty(line.Trim()))
                        continue;

                    if (Regex.IsMatch(line, @"On\s+\w+\s+GoTo", RegexOptions.IgnoreCase) && !bOnErrorGoTo)
                    {
                        // formalize Error name
                        string errName = Regex.Match(line, @"On\s+(?<err>\w+)\s+GoTo", RegexOptions.IgnoreCase).Groups["err"].ToString().Trim();
                        if (!errName.Equals("Error", StringComparison.OrdinalIgnoreCase))
                            procedure.ModifiedErrorName = true;
                        lines[j] = lines[j].Replace(" " + errName + " ", " Error ");

                        // 1st time disappear On Error GoTo
                        if (!bOnErrorGoTo)
                        {
                            firstOnErrorGoToIndex = tmpResult.FindIndex(p => p.IndexOf(line) != -1);
                            tmpResult.RemoveAt(firstOnErrorGoToIndex);
                            tmpResult.Insert(firstOnErrorGoToIndex, lines[j]);
                            procedure.ExistOnErrorGoTo = true;
                            onErrorLine = lines[j];

                            label = Regex.Match(lines[j], @"On Error GoTo\s+(?<label>\w+)$", RegexOptions.IgnoreCase).Groups["label"].ToString().Trim();
                            bOnErrorGoTo = true;
                        }
                        else
                        {
                            int idx = tmpResult.FindIndex(firstOnErrorGoToIndex + 1, p => p.IndexOf(line) != -1);
                            tmpResult.RemoveAt(idx);
                            tmpResult.Insert(idx, "\'" + lines[j]);
                        }
                    }

                    if (Regex.IsMatch(line, @OnErrorResumeNext, RegexOptions.IgnoreCase))
                        procedure.ExistOnErrorResumeNext = true;

                    if (Regex.IsMatch(line, @"\s*(^Stop|[^\w]Stop)([^\w]|$)", RegexOptions.IgnoreCase))
                        procedure.StopKeyword = true;
                    if (Regex.IsMatch(line, @"\s*([^\w]MsgBox|^MsgBox)[\s\(]", RegexOptions.IgnoreCase))
                        procedure.MsgBoxKeyword = true;

                    if (((line.Contains(":") && line.IndexOf("handler", StringComparison.OrdinalIgnoreCase) != -1) || line.Contains(label + ":")) && bExitFunction)
                    {
                        errHandlerList.Add(exitFunctionLine);
                        procedure.ExistErrHandler = true;
                    }

                    if (line.IndexOf("END FUNCTION", StringComparison.OrdinalIgnoreCase) != -1 ||
                        line.IndexOf("END SUB", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        i = j;

                        // this must put before below judgement line
                        originalProcedures.Add(procedure.copyObject());

                        // Exit Function only
                        if (bExitFunction && !procedure.ExistErrHandler)
                        {
                            int idx = tmpResult.FindIndex(p => p.Equals(exitFunctionLine));
                            tmpResult.RemoveAt(idx);
                            tmpResult.Insert(idx, "\'" + exitFunctionLine);
                        }

                        // head only, no tail
                        if (procedure.ExistOnErrorGoTo && !procedure.ExistErrHandler)
                        {
                            int idx = tmpResult.FindIndex(p => p.Equals(onErrorLine));
                            tmpResult.RemoveAt(idx);
                            tmpResult.Insert(idx, "\'" + onErrorLine);
                            procedure.ExistOnErrorGoTo = false;

                        }

                        // no head, tail only
                        if (!procedure.ExistOnErrorGoTo && procedure.ExistErrHandler)
                        {
                            foreach(string item in errHandlerList)
                            {
                                if (string.IsNullOrEmpty(item.Trim()))
                                    continue;
                                int idx = tmpResult.FindIndex(p => p.Equals(item));
                                tmpResult.RemoveAt(idx);
                                tmpResult.Insert(idx, "\'" + item);
                            }
                            procedure.ExistErrHandler = false;
                        }

                        // empty error handler content
                        if (procedure.ExistErrHandler && errHandlerList.Count < 3)
                        {
                            foreach(string item in errHandlerList)
                            {
                                if (string.IsNullOrEmpty(item.Trim()))
                                    continue;
                                int idx = tmpResult.FindIndex(p => p.Equals(item));
                                tmpResult.RemoveAt(idx);
                                tmpResult.Insert(idx, "\'" + item);
                            }
                            procedure.ExistErrHandler = false;
                        }

                        procedures.Add(procedure);
                        break;
                    }

                    if (procedure.ExistErrHandler)
                        errHandlerList.Add(lines[j]);

                    // this judgement should be put at last step
                    if (Regex.IsMatch(line, @"Exit Function|Exit Sub", RegexOptions.IgnoreCase) && !Regex.IsMatch(line, @"[\s]if\s+.*then", RegexOptions.IgnoreCase))
                    {
                        exitFunctionLine = lines[j];
                        bExitFunction = true;
                    }
                    //else
                    //{
                    //    bExitFunction = false;
                    //}
                }
                result.AddRange(tmpResult);
            }

            return result;
        }

        private void WriteUILog(string message, Color color)
        {
            if (_appendText == null)
                return;
            _appendText.Invoke(message, color);
        }
        private void WriteReport()
        {
            WriteUILog("Starting to print report ...", Color.Black);
            using (var ep = new ExcelPackage(new FileInfo(_outputFilePath)))
            {
                ExcelWorksheet ws = ep.Workbook.Worksheets.Add("ErrorHandler");
                ws.Cells[1, 1].LoadFromCollection(_result, true);
                for (var i = 1; i < 6; i++)
                {
                    ws.Cells[1, i].Style.Font.Bold = true;
                    ws.Cells[1, i].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);
                }           
                for (int i = 0; i < _result.Count; ++i)
                {
                    if (_result[i].Modified == true)
                        ws.Cells[i + 2, 3].Style.Font.Color.SetColor(Color.Red);
                    if (_result[i].StopKeyword == true)
                        ws.Cells[i + 2, 4].Style.Font.Color.SetColor(Color.Red);
                    if (_result[i].MsgBoxKeyword == true)
                        ws.Cells[i + 2, 5].Style.Font.Color.SetColor(Color.Red);
                }
                // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add start
                ws.Cells["A1:E1"].AutoFilter = true;
                // 2022-02-11  Steven Chen    #310	          Simplify Output Option ErrorHandler sheet(Remove Pass Item) add end
                ws.Cells.AutoFitColumns();
                // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg start
                //ws = ep.Workbook.Worksheets.Add("Option Explicit");
                //ws.Cells[1, 1].LoadFromCollection(_oeResult, true);

                //for (int i = 1; i < 3; ++i)
                //{
                ws = ep.Workbook.Worksheets.Add("Option Explicit Check");
                ws.Cells[1, 1].LoadFromCollection(_oeResult, true);
                ws.Cells[1, 1].Value = "The Files Which Option Explicit Not Exist";

                for (int i = 1; i < 2; ++i)
                {
                // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg end
                    ws.Cells[1, i].Style.Font.Bold = true;
                    ws.Cells[1, i].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);
                }
                ws.Cells.AutoFitColumns();

                if (!Directory.Exists(_outputFolder))
                    Directory.CreateDirectory(_outputFolder);
                ep.Save();
            }
            WriteUILog("Save Path: "+ _outputFilePath, Color.DarkBlue);
        }
        
    }
    
}
