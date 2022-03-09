using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Library.Common;

namespace Library.Input
{
    public class TestProgramReader
    {
        public TestProgramReader()
        {
             
        }

        public ExcelWorkbook LoadAutoGenTestProgram(string testProgramPath)
        {
            if (testProgramPath.EndsWith(".igxl"))
            {
                string excelfile = testProgramPath.Replace(".igxl", ".xlsm");
                ExportWorkBook(testProgramPath, testProgramPath.Replace(".igxl", ".xlsm"));
                var finfoDum = new FileInfo(excelfile);
                ExcelPackage package = new ExcelPackage(finfoDum);
                return package.Workbook;
            }
            else
            {
                FileInfo fileInfo = new FileInfo(testProgramPath);
                var dummyExcel = Regex.Replace(fileInfo.FullName, fileInfo.Extension, "_DUM" + fileInfo.Extension);
                fileInfo.CopyTo(dummyExcel, true);
                var finfoDum = new FileInfo(dummyExcel);
                ExcelPackage package = new ExcelPackage(finfoDum);
                File.Delete(dummyExcel);
                return package.Workbook;
            }
        }

        public ExcelPackage LoadProductionTestProgram(string testProgramPath, string currentTimeStr)
        {
            string excelfile = testProgramPath.Replace(".igxl", ".xlsm");
            if (testProgramPath.EndsWith(".igxl"))
            {
                ExportWorkBook(testProgramPath, excelfile);
            }

            FileInfo fileInfo = new FileInfo(excelfile);
            var dummyExcel = CommonData.GetInstance().OutputPath + "\\" + Regex.Replace(fileInfo.Name, fileInfo.Extension, "_Log_" + currentTimeStr + fileInfo.Extension);
            fileInfo.CopyTo(dummyExcel, true);
            var finfoDum = new FileInfo(dummyExcel);
            ExcelPackage package = new ExcelPackage(finfoDum);

            return package;
        }

        public void ExportWorkBook(string v900testprogrampath, string exceltestprogrampath)
        {
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
            if (File.Exists(exportWorkBookCmd))
            {
                string option = "-w \"" + v900testprogrampath + "\" -e \"" + exceltestprogrampath + "\"";
                RunCmd(exportWorkBookCmd, option);
            }
        }

        private void RunCmd(string cmd, string argment = "")
        {
            Process nProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = cmd;
            startInfo.Arguments = argment;
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        public static bool DeleteFolder(string folderPath)
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.Arguments = "/c " + "rd \"" + folderPath + "\" /S /Q";
            p.Start();

            p.WaitForExit();
            p.Close();
            return true;
        }
    }
}
