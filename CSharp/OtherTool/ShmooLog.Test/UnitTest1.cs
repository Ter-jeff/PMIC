using CommonLib.Extension;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace ShmooLog.Test
{
    [TestClass]
    public class UnitTest1
    {
        private const string Input = "Input";
        private const string Output = "Output";
        private const string Expected = "Expected";

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int id);

        [TestMethod]
        [DeploymentItem(Input + @"\QF25_20220129_AI_21_CP1_x10y7_Shmoo.txt", Input)]
        [DeploymentItem(Expected, Expected)]
        public void TestMethod1()
        {
            var log = Path.Combine(Directory.GetCurrentDirectory(),
                Input + @"\QF25_20220129_AI_21_CP1_x10y7_Shmoo.txt");
            var expectedPath = Path.Combine(Directory.GetCurrentDirectory(), Expected);
            var outputPath = Path.Combine(Directory.GetCurrentDirectory(), Output);
            var shmooLog = new ShmooLog(log);
            shmooLog.ParseEachDevices();
            var shmooLogs = new ShmooLogs { shmooLog };
            var report = shmooLogs.ConvertExcel(outputPath);

            var xlApp = new Application { DisplayAlerts = false, UserControl = false };
            var workbook = xlApp.Workbooks.Open(report, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (!Directory.Exists(outputPath))
                Directory.CreateDirectory(outputPath);
            foreach (Worksheet sheet in workbook.Worksheets) sheet.ExportTxt(outputPath);

            workbook.Close(false);
            var intPtr = new IntPtr(xlApp.Hwnd);
            int excelProcessId;
            GetWindowThreadProcessId(intPtr, out excelProcessId);
            var excelProcess = Process.GetProcessById(excelProcessId);
            if (excelProcess != null)
            {
                excelProcess.Kill();
                excelProcess.Dispose();
            }

            var isFail = false;
            var files = Directory.GetFiles(expectedPath, "*.*", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                var outputFile = Path.Combine(Path.GetDirectoryName(file).Replace(Expected, Output),
                    Path.GetFileName(file));
                if (File.Exists(outputFile))
                {
                    if (!TxtCompere(file, outputFile)) // text diff
                                                       //if (!FileCompare(file, outputFile)) binary diff
                    {
                        isFail = true;
                        Debug.Write(outputFile + "is not as expected file !!!");
                    }
                }
                else
                {
                    isFail = true;
                    Debug.Write("Can not find file !!!");
                }
            }

            if (isFail)
                Assert.Fail();
        }

        private bool TxtCompere(string inputFile, string outputFile)
        {
            var inputs = File.ReadAllLines(inputFile).ToList();
            var outputs = File.ReadAllLines(outputFile).ToList();

            if (inputs.Count != outputs.Count) return false;

            for (var index = 0; index < outputs.Count; index++)
                if (inputs[index] != outputs[index])
                    return false;
            return true;
        }

        private bool FileCompare(string file1, string file2)
        {
            int file1byte;
            int file2byte;
            FileStream fs1;
            FileStream fs2;

            // Determine if the same file was referenced two times.
            if (file1 == file2)
                // Return true to indicate that the files are the same.
                return true;

            // Open the two files.
            fs1 = new FileStream(file1, FileMode.Open);
            fs2 = new FileStream(file2, FileMode.Open);

            // Check the file sizes. If they are not the same, the files
            // are not the same.
            if (fs1.Length != fs2.Length)
            {
                // Close the file
                fs1.Close();
                fs2.Close();

                // Return false to indicate files are different
                return false;
            }

            // Read and compare a byte from each file until either a
            // non-matching set of bytes is found or until the end of
            // file1 is reached.
            do
            {
                // Read one byte from each file.
                file1byte = fs1.ReadByte();
                file2byte = fs2.ReadByte();
            } while (file1byte == file2byte && file1byte != -1);

            // Close the files.
            fs1.Close();
            fs2.Close();

            // Return the success of the comparison. "file1byte" is
            // equal to "file2byte" at this point only if the files are
            // the same.
            return file1byte - file2byte == 0;
        }
    }
}