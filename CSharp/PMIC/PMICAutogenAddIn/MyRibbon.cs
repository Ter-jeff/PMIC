using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Utility;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using OfficeOpenXml;
using PmicAutogen;
using PmicAutogen.Inputs.TestPlan;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace PMICAutogenAddIn
{
    public partial class MyRibbon
    {
        private void button_Validate_Click(object sender, RibbonControlEventArgs e)
        {
            var startTime = TimeProvider.Current.Now;
            try
            {
                SetCursorToWaiting();
                Globals.ThisAddIn.Application.StatusBar = "Starting to validate ...";
                ErrorManager.Initialize();
                var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                SheetStructureManager.Initialize();

                Globals.ThisAddIn.Application.StatusBar = "Saving & Copying file ...";
                workBook.Save();
                var extension = Path.GetExtension(workBook.FullName);
                var targetFile = Path.GetFileName(workBook.FullName).Replace(extension, "_Temp" + extension);
                if (File.Exists(targetFile))
                    File.Delete(targetFile);
                if (workBook.FullName != null)
                    File.Copy(workBook.FullName, targetFile);

                using (var excel = new ExcelPackage(new FileInfo(targetFile)))
                {
                    var sheetCheckManager = new TestPlanManager();
                    sheetCheckManager.CheckAll(excel.Workbook, Globals.ThisAddIn.Application);
                }
                File.Delete(targetFile);

                var errors = ErrorManager.GetErrors();
                if (errors.Count() > 0)
                {
                    ErrorManager.GenErrorReportVSTO(workBook, "ErrorReport");
                    var errorCut = errors.Count(x => x.ErrorLevel == EnumErrorLevel.Error);
                    var warningCut = errors.Count(x => x.ErrorLevel == EnumErrorLevel.Warning);
                    Globals.ThisAddIn.Application.StatusBar = string.Format("Completed ... ({0})", (TimeProvider.Current.Now - startTime).ToString(@"hh\:mm\:ss"));
                    MessageBox.Show(String.Format("Totol Errors : {0} \r\nTotol Warnings : {1}", errorCut, warningCut),
                        "Find Errors !!!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    Globals.ThisAddIn.Application.StatusBar = string.Format("Process time : {0}", (TimeProvider.Current.Now - startTime).ToString(@"hh\:mm\:ss"));
                    MessageBox.Show("Validation is done !!!");
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString(), "", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetCursorToDefault();
            }
        }

        private void button_Autogen_Click(object sender, RibbonControlEventArgs e)
        {
            if (CheckEnvironment() == false)
            {
                MessageBox.Show("No IGXL found.");
                return;
            }

            //var pmicMainForm = new PmicMainForm(Globals.ThisAddIn.Application.ActiveWorkbook);
            //pmicMainForm.Show();
            var pmicMainForm = new PmicMainWindow(Globals.ThisAddIn.Application.ActiveWorkbook);
            pmicMainForm.ShowDialog();
        }

        private void button_Help_Click(object sender, RibbonControlEventArgs e)
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //var resourceNames = assembly.GetManifestResourceNames();
            //foreach (var resourceName in resourceNames)
            //    if (resourceName.EndsWith(".PMICAutogenHelp.chm", StringComparison.CurrentCultureIgnoreCase))
            //    {
            //        var stream = assembly.GetManifestResourceStream(resourceName);
            //        var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //        var helpFile = Path.Combine(Path.GetDirectoryName(workBook.FullName), "PMICAutogenHelp.chm");
            //        var fileStream = File.Create(helpFile);
            //        stream.Seek(0, SeekOrigin.Begin);
            //        stream.CopyTo(fileStream);
            //        fileStream.Close();
            //        if (File.Exists(helpFile))
            //            Process.Start(helpFile);
            //        break;
            //    }
            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase)
                .Replace("file:\\", "");
            var helpPath = Path.Combine(exePath, "Help\\PMICAutogenHelp.chm");
            if (File.Exists(helpPath))
            {
                var startInfo = new ProcessStartInfo(helpPath);
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start(startInfo);
            }
        }

        private void button_History_Click(object sender, RibbonControlEventArgs e)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith(".PMICAutoGenAddin_Release_Notes.docx",
                        StringComparison.CurrentCultureIgnoreCase))
                {
                    var stream = assembly.GetManifestResourceStream(resourceName);
                    var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                    var docx = Path.Combine(Path.GetDirectoryName(workBook.FullName),
                        "PMICAutoGenAddin_Release_Notes.docx");
                    var fileStream = File.Create(docx);
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                    fileStream.Close();
                    if (File.Exists(docx))
                        Process.Start(docx);
                    break;
                }
        }

        private void button_Back_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.GoBack();
        }

        private bool CheckEnvironment()
        {
            var igxlRoot = Environment.GetEnvironmentVariable("IGXLROOT");
            if (string.IsNullOrEmpty(igxlRoot))
                return false;
            return true;
        }

        private static void SetCursorToWaiting()
        {
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = XlMousePointer.xlWait;
        }

        private static void SetCursorToDefault()
        {
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = XlMousePointer.xlDefault;
        }
    }
}