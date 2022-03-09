using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Library.Function
{
    public class UnitTestResult
    {
        public string Expected { get; set; }
        public string Output { get; set; }
        public string Link { get; set; }
        public string Remark { get; set; }
    }

    public class FileComparision
    {
        private readonly string _expectedPath;
        private readonly string _outputPath;
        private readonly string _batFilePath;
        private readonly string _winMergeFile;
        private readonly string _report;

        public readonly List<UnitTestResult> UnitTestResults = new List<UnitTestResult>();

        public FileComparision(string expectedPath, string outputPath, string report, string batFilePath,
            string winMergeFile)
        {
            _expectedPath = expectedPath;
            _outputPath = outputPath;
            _batFilePath = batFilePath;
            _winMergeFile = winMergeFile;
            _report = report;
        }

        public bool IsFolderCompareFail()
        {
            List<string> expectedFiles = Directory.GetFiles(_expectedPath, "*", SearchOption.AllDirectories).ToList();
            expectedFiles = ExcludingFile(expectedFiles, "xls*");
            List<string> outputFiles = Directory.GetFiles(_outputPath, "*", SearchOption.AllDirectories).ToList();
            outputFiles = ExcludingFile(outputFiles, "xls*");
            bool flag = false;

            foreach (string expectedFile in expectedFiles)
            {
                foreach (string outputFile in outputFiles)
                {
                    if (Path.GetFileName(expectedFile) == Path.GetFileName(outputFile))
                    {
                        if (!FileCompare(expectedFile, outputFile))
                        {
                            UnitTestResult unitTestResult = new UnitTestResult();
                            unitTestResult.Expected = Path.GetFileName(expectedFile);
                            unitTestResult.Output = Path.GetFileName(outputFile);
                            string changeExtension = Path.ChangeExtension(outputFile, ".bat");
                            if (changeExtension != null)
                            {
                                string batFile = Path.Combine(_batFilePath,
                                    Path.GetFileName(changeExtension));
                                unitTestResult.Link = $"HYPERLINK(\"{batFile}\", " +
                                                      "\"Link\")";
                                UnitTestResults.Add(unitTestResult);
                                CreateBatFile(expectedFile, outputFile, batFile);
                            }
                        }

                        flag = true;
                        break;
                    }
                }

                if (!flag)
                {
                    UnitTestResult unitTestResult = new UnitTestResult();
                    unitTestResult.Expected = Path.GetFileName(expectedFile);
                    unitTestResult.Output = "";
                    unitTestResult.Remark = "Missing output file !!!";
                    UnitTestResults.Add(unitTestResult);
                }
            }

            List<string> misses = outputFiles
                .Where(x => expectedFiles.All(y => Path.GetFileName(x) != Path.GetFileName(y)))
                .ToList();
            foreach (string miss in misses)
            {
                UnitTestResult unitTestResult = new UnitTestResult();
                unitTestResult.Expected = "";
                unitTestResult.Output = Path.GetFileName(miss);
                unitTestResult.Remark = "Missing expected file !!!";
                UnitTestResults.Add(unitTestResult);
            }

            WriteDiff();
            return UnitTestResults.Count != 0;
        }

        private void CreateBatFile(string expectedFile, string outputFile, string batFile)
        {
            List<string> lines = new List<string>();
            lines.Add("@echo off");
            lines.Add("start " + _winMergeFile + " " +
                      expectedFile + " " + outputFile);
            if (outputFile != null)
            {
                File.WriteAllLines(batFile, lines);
            }
        }

        public bool FileCompare(string file1, string file2)
        {
            int file1Byte;
            int file2Byte;

            if (file1 == file2)
            {
                return true;
            }

            FileStream fs1 = new FileStream(file1, FileMode.Open);
            FileStream fs2 = new FileStream(file2, FileMode.Open);

            if (fs1.Length != fs2.Length)
            {
                fs1.Close();
                fs2.Close();
                return false;
            }

            do
            {
                file1Byte = fs1.ReadByte();
                file2Byte = fs2.ReadByte();
            } while (file1Byte == file2Byte && file1Byte != -1);

            fs1.Close();
            fs2.Close();

            return file1Byte - file2Byte == 0;
        }

        private void WriteDiff()
        {
            if (UnitTestResults.Count == 0)
            {
                return;
            }

            if (Epplus.IsExcelOpened(_report))
            {
                if (MessageBox.Show("The files was opened,and please close all excel workbooks !!!", "Excel closing",
                        MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Epplus.KillExcel();
                }
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(_report)))
            {
                ExcelWorkbook workbook = package.Workbook;
                string sheetName = "FileDiff";
                workbook.DeleteSheet(sheetName);
                ExcelWorksheet workSheet = workbook.AddSheet(sheetName);
                workSheet.Cells[1, 1].LoadFromCollection(UnitTestResults, true);
                workSheet.SetFormula(3);
                workSheet.SetHeaderStyle();
                workSheet.SetHairBorder();
                workSheet.Cells.AutoFitColumns();
                package.Save();
            }

            Process.Start(_report);
        }

        public List<string> ExcludingFile(List<string> files, string extension)
        {
            extension = extension.Replace("*", "\\w?");
            List<string> excludingFiles = new List<string>();
            foreach (string file in files)
            {
                if (file != null && Regex.IsMatch(Path.GetExtension(file), extension))
                {
                    continue;
                }

                excludingFiles.Add(file);
            }

            return excludingFiles;
        }
    }
}