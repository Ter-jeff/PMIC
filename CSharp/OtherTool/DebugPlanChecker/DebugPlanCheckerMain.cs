using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using CommonReaderLib.DebugPlan;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace DebugPlanChecker
{
    internal class DebugPlanCheckerMain
    {
        private readonly string _debugPlan;
        private readonly string _patternPath;
        private readonly string _testProgram;
        private readonly string _outputFolder;

        public DebugPlanCheckerMain(string debugPlan, string patternPath, string testProgram, string outputFolder)
        {
            _debugPlan = debugPlan;
            _patternPath = patternPath;
            _testProgram = testProgram;
            _outputFolder = outputFolder;
        }

        internal void WorkFlow()
        {
            try
            {
                if (_debugPlan.IsOpened())
                {
                    var result = MessageBox.Show("Please close " + _debugPlan + " !!!", "", MessageBoxButton.OKCancel
                           , MessageBoxImage.Error);
                    if (result == MessageBoxResult.Cancel) return;
                }
                Response.Report("Starting to read Debug Test Plan ...", EnumMessageLevel.General, 0);
                var debugTestPlan = new DebugPlanMain(_debugPlan);
                debugTestPlan.Read();
                Response.Report("Starting to check Debug Test Plan ...", EnumMessageLevel.General, 50);
                debugTestPlan.CheckAll(_patternPath, Path.Combine(_patternPath, "TimeSet"), _testProgram);

                if (debugTestPlan.Errors.Count > 0)
                {
                    var fileName = Path.GetFileNameWithoutExtension(_debugPlan) + "_Report_" + TimeProvider.Current.Now.ToString("yyyyMMddhhmmss")
                        + Path.GetExtension(_debugPlan);
                    var outputReport = Path.Combine(_outputFolder, fileName);
                    File.Copy(_debugPlan, outputReport, true);
                    Response.Report("Total Error Count : " + debugTestPlan.Errors.Count, EnumMessageLevel.Error, 80);
                    ErrorManager.AddErrors(debugTestPlan.Errors);
                    var app = new Microsoft.Office.Interop.Excel.Application();
                    app.DisplayAlerts = false;
                    var workbook = app.Workbooks.Open(outputReport, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing);

                    try
                    {
                        ErrorManager.GenErrorReport(workbook, "ErrorReport");
                        Response.Report("Error report : " + workbook.FullName + "...", EnumMessageLevel.Error, 80);
                    }
                    finally
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                        Marshal.FinalReleaseComObject(workbook);
                        Marshal.FinalReleaseComObject(app.Workbooks);
                        app.Quit();
                        Marshal.FinalReleaseComObject(app);
                    }
                }
                Response.Report("Finishing to check Debug Test Plan ...", EnumMessageLevel.General, 100);
            }
            catch (Exception ex)
            {
                Response.Report(ex.StackTrace.ToString(), EnumMessageLevel.Error, 0);
            }

        }
    }
}