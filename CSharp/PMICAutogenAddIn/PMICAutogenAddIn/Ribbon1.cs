using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using AutomationCommon.EpplusErrorReport;
using Microsoft.Office.Tools.Ribbon;
using OfficeOpenXml;
using PmicAutogen;
using PmicAutogen.Inputs.TestPlan;

namespace PMICAutogenAddIn
{
    public partial class Ribbon1
    {
        private void button_Validate_Click(object sender, RibbonControlEventArgs e)
        {
            var startTime = DateTime.Now;
            Globals.ThisAddIn.Application.StatusBar = "Starting to validate ...";
            EpplusErrorManager.Initialize();
            var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            SheetStructureManager.Initialize();

            Globals.ThisAddIn.Application.StatusBar = "Saving & Copying file ...";
            workBook.Save();
            var extension = Path.GetExtension(workBook.FullName);
            var targetFile = Path.GetFileName(workBook.FullName).Replace(extension, "_Temp" + extension);
            if (File.Exists(targetFile))
                File.Delete(targetFile);
            if (workBook.FullName != null) File.Copy(workBook.FullName, targetFile);


            using (var excel = new ExcelPackage(new FileInfo(targetFile)))
            {
                var sheetCheckManager = new TestPlanManager();
                sheetCheckManager.CheckAll(excel.Workbook, Globals.ThisAddIn.Application);
            }

            File.Delete(targetFile);
            EpplusErrorManager.GenErrorReport(workBook, "ErrorReport");

           // button_Autogen.Enabled = EpplusErrorManager.GetErrorCount() == 0;
            Globals.ThisAddIn.Application.StatusBar =
                string.Format("Completed ... ({0:hh\\:mm\\:ss})", DateTime.Now - startTime);
        }

        private void button_Autogen_Click(object sender, RibbonControlEventArgs e)
        {
            if (CheckEnvironment() == false)
            {
                System.Windows.Forms.MessageBox.Show("No IGXL found.");
                return;
            }
            var pmicMainForm = new PmicMainForm(Globals.ThisAddIn.Application.ActiveWorkbook);
            pmicMainForm.Show();
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
            string exePath = Assembly.GetExecutingAssembly().GetName().CodeBase;
            string helpPath = Path.Combine(Path.GetDirectoryName(exePath), "Help\\PMICAutogenHelp.chm");
            helpPath = helpPath.Replace("file:\\", "");//new System.Uri(helpPath).LocalPath;
            if (File.Exists(helpPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(helpPath);
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start(startInfo);
            }
        }

        private bool CheckEnvironment()
        {
            string igxlRoot = Environment.GetEnvironmentVariable("IGXLROOT");
            if (string.IsNullOrEmpty(igxlRoot))
                return false;
            else
                return true;
        }
    }
}