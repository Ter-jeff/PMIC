using AutoIgxl.Reader;
using CommonLib.Enum;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace AutoIgxl
{
    public class AutoIgxlMain
    {
        private readonly string _outputIgxl;
        public Logger Logger = LogManager.GetCurrentClassLogger();

        public AutoIgxlMain(string outputIgxl)
        {
            _outputIgxl = outputIgxl;
        }

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        public Application GetIgxl()
        {
            //excel classID {00000304-0000-0000-c000-000000000046}
            IntPtr numFetched = new IntPtr();
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            var monikers = new IMoniker[1];
            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);
                Guid classId;
                monikers[0].GetClassID(out classId);
                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);
                if (runningObjectVal is Application)
                {
                    var excel = (Application)runningObjectVal;
                    //if (excel.Caption.ToLower().Contains(".igxl"))
                    return excel;
                }
            }

            return null;
        }

        public void RunProgram(RunCondition runCondition)
        {
            try
            {
                if (HasExcel())
                {
                    var result =
                        MessageBox.Show(
                            @"Please close all excel and Igxl !!!\nClick Ok will continue \nClick cancel will abort !!! ",
                            "", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    Logger.Trace("Need to close all excel and Igxl !!!");
                    if (result == MessageBoxResult.Cancel)
                        return;
                }
                //KillExcel();

                if (runCondition == null) return;

                using (var myProcess = Process.Start(_outputIgxl))
                {
                    while (myProcess != null && !myProcess.HasExited)
                    {
                    }

                    var excelApp = GetIgxl();
                    excelApp.Visible = true;

                    Logger.Trace("Starting to validate test program ...");
                    excelApp.Application.Run("tl_dt_AValidate");

                    ErrorSheet errorSheet = null;
                    foreach (Worksheet sheet in excelApp.ActiveWorkbook.Worksheets)
                        if (sheet.Name.StartsWith("-Errors-", StringComparison.CurrentCultureIgnoreCase))
                        {
                            var errorSheetReader = new ErrorSheetReader();
                            errorSheet = errorSheetReader.ReadSheet(sheet);
                            foreach (var row in errorSheet.Rows)
                                Logger.Error("[" + EnumNLogMessage.Input + "] " + "SheetName:" + row.SheetName +
                                             "ErrorCell:" + row.Cell + ", ErrorCode:" +
                                             row.ErrorCode + "ErrorMessage:" + row.ErrorMessage);
                        }

                    if (errorSheet == null)
                    {
                        AddSetMacro(excelApp);
                        if (runCondition.ExecSites != null && runCondition.ExecSites.Any())
                            foreach (var totalSite in runCondition.TotalSites)
                            {
                                var flag = runCondition.ExecSites.Exists(
                                    x => x.Equals(totalSite, StringComparison.CurrentCultureIgnoreCase));
                                var siteNumber =
                                    int.Parse(Regex.Replace(totalSite, "^Site", "", RegexOptions.IgnoreCase));
                                excelApp.Application.Run(flag ? "tl_ExecEnableSite" : "tl_ExecDisableSite", siteNumber);
                            }
                        if (!string.IsNullOrEmpty(runCondition.LotId))
                            excelApp.Application.Run("SetLotID", runCondition.LotId);
                        if (!string.IsNullOrEmpty(runCondition.WaferId))
                            excelApp.Application.Run("SetWaferID", runCondition.WaferId);
                        if (!string.IsNullOrEmpty(runCondition.OutputLog))
                            excelApp.Application.Run("SetOutputTxt", runCondition.OutputLog);
                        if (!string.IsNullOrEmpty(runCondition.Job))
                            excelApp.Application.Run("SetCurrentJob", runCondition.Job);
                        if (!string.IsNullOrEmpty(runCondition.DoAll.ToString()))
                            excelApp.Application.Run("SetDoAll", runCondition.DoAll.ToString());
                        if (!string.IsNullOrEmpty(runCondition.OverrideFailStop.ToString()))
                            excelApp.Application.Run("SetOverrideFailStop", runCondition.OverrideFailStop.ToString());
                        if (!string.IsNullOrEmpty(runCondition.SetXy) && runCondition.SetXy.Contains(','))
                        {
                            var xy = runCondition.SetXy.Split(',');
                            excelApp.Application.Run("SetXY", int.Parse(xy.First()), int.Parse(xy.Last()));
                        }

                        if (runCondition.ExecEnableWords != null)
                            foreach (var enableWord in runCondition.TotalEnableWords)
                            {
                                var flag = runCondition.ExecEnableWords.Exists(
                                    x => x.Equals(enableWord, StringComparison.CurrentCultureIgnoreCase));
                                excelApp.Application.Run("tl_ExecSetEnableWord", enableWord, flag);
                            }

                        //RemoveSetMacro(excelApp);
                        Logger.Trace("Starting to run test program ...");
                        excelApp.Application.Run("tl_ProgramRun");
                        if (File.Exists(runCondition.OutputLog))
                        {
                            Logger.Trace("Starting to unload test program ...");
                            excelApp.Application.Run("tl_ProgramUnload");
                        }

                        excelApp.DisplayAlerts = false;
                        object misValue = Missing.Value;
                        if (File.Exists(runCondition.OutputLog))
                            Logger.Trace("Output log => " + runCondition.OutputLog + " ...");
                        else
                            Logger.Error("[" + EnumNLogMessage.Input + "] " + "No output log !!!");
                        excelApp.ActiveWorkbook.Close(false, misValue, misValue);
                        excelApp.Quit();
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error(e.StackTrace);
            }
        }

        private bool HasExcel()
        {
            try
            {
                var excel = Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private void KillExcel()
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                FileName = "taskkill",
                Arguments = "/f /im excel.exe"
            };
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        public void RemoveSetMacro(Application excelApp)
        {
            const string vbtTemp = "VBT_Temp";
            var workbook = excelApp.ActiveWorkbook;
            foreach (VBComponent vbComponent in workbook.VBProject.VBComponents)
                if (vbComponent.CodeModule.Name == vbtTemp)
                {
                    workbook.VBProject.VBComponents.Remove(vbComponent);
                    return;
                }
        }

        public void AddSetMacro(Application excelApp)
        {
            const string vbtTemp = "VBT_Temp";
            var workbook = excelApp.ActiveWorkbook;
            var newStandardModule = GetVbComponents(workbook, vbtTemp);
            var codeModule = newStandardModule.CodeModule;
            codeModule.DeleteLines(1, codeModule.CountOfLines);
            var codeText = SetLotId();
            codeText += SetWaferId();
            codeText += SetOutputTxt();
            codeText += SetCurrentJob();
            codeText += SetDoAll();
            codeText += SetSetOverrideFailStop();
            codeModule.Name = vbtTemp;
            codeModule.InsertLines(1, codeText);
        }

        private string SetLotId()
        {
            var codeText = "Public Sub SetLotID(Name As String)" + "\r\n";
            codeText += "  TheExec.Datalog.Setup.LotSetup.LotID = Name" + "\r\n";
            codeText += "  TheExec.Datalog.ApplySetup" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string SetWaferId()
        {
            var codeText = "Public Sub SetWaferID(Name As String)" + "\r\n";
            codeText += "  TheExec.Datalog.Setup.WaferSetup.ID = Name" + "\r\n";
            codeText += "  TheExec.Datalog.ApplySetup" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string SetOutputTxt()
        {
            var codeText = "Public Sub SetOutputTxt(Name As String)" + "\r\n";
            codeText += "  TheExec.Datalog.Setup.DatalogSetup.TextOutput = True" + "\r\n";
            codeText += "  TheExec.Datalog.Setup.DatalogSetup.TextOutputFile = Name" + "\r\n";
            codeText += "  TheExec.Datalog.Setup.DatalogSetup.DatalogOn = True" + "\r\n";
            codeText += "  TheExec.Datalog.ApplySetup" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string SetCurrentJob()
        {
            var codeText = "Public Sub SetCurrentJob(Name As String)" + "\r\n";
            codeText += "  TheExec.CurrentJob = Name" + "\r\n";
            codeText += "  TheExec.Datalog.ApplySetup" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string SetDoAll()
        {
            var codeText = "Public Sub SetDoAll(Flag As Boolean)" + "\r\n";
            codeText += "  TheExec.RunOptions.DoAll = Flag" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private string SetSetOverrideFailStop()
        {
            var codeText = "Public Sub SetOverrideFailStop(Flag As Boolean)" + "\r\n";
            codeText += "  TheExec.RunOptions.OverrideFailStop = Flag" + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        private VBComponent GetVbComponents(Workbook workbook, string vbtName)
        {
            foreach (VBComponent vbComponent in workbook.VBProject.VBComponents)
                if (vbComponent.CodeModule.Name == vbtName)
                    return vbComponent;
            return workbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
        }
    }
}