using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Test.FileDiff
{
    internal class FileComparisonReport
    {
        private const string ReposrtName = "FileDiffReport";
        private const string Output = "Output";
        private const string Expected = "Expected";
        private readonly string _kdiff3 = Directory.GetCurrentDirectory() + @"\KDiff3\kdiff3";

        internal bool Main(string tarDir, string expected, bool isReport)
        {
            var fail = false;
            var filter = ".txt|.xml|.pa|.bas|.cls";
            var tar = Directory.GetFiles(tarDir, "*.*", SearchOption.AllDirectories)
                .Where(x => Regex.IsMatch(Path.GetExtension(x), filter, RegexOptions.IgnoreCase))
                .ToDictionary(x => x, x => x.Replace(tarDir, ""));
            var exp = Directory.GetFiles(expected, "*.*", SearchOption.AllDirectories)
                .Where(x => Regex.IsMatch(Path.GetExtension(x), filter, RegexOptions.IgnoreCase))
                .ToDictionary(x => x, x => x.Replace(expected, ""));
            var add = tar.Values.Except(exp.Values).ToList();
            var missing = exp.Values.Except(tar.Values).ToList();
            var find = tar.Values.Intersect(exp.Values).ToList();
            var fileComparer = new FileComparer();
            fileComparer.AddItems = add;
            fileComparer.MissingItems = missing;
            fileComparer.Output = tarDir;
            fileComparer.Expected = expected;
            foreach (var item in find)
                if (!TexCompare(tarDir + item, expected + item))
                {
                    fileComparer.DiffItems.Add(
                        new FileDiff { OutputFile = tarDir + item, ExpectedFile = expected + item });
                    fail = true;
                }

            if (isReport)
                GenerateReport(fileComparer);
            return fail;
        }

        private bool TexCompare(string oldFile, string newFile)
        {
            var oldLines = File.ReadAllLines(oldFile);
            var newLines = File.ReadAllLines(newFile);
            if (oldLines.Count() != newLines.Count()) return false;
            for (var i = 0; i < oldLines.Count(); i++)
                if (oldLines[i] != newLines[i])
                    return false;
            return true;
        }

        private void GenerateReport(FileComparer fileComparer)
        {
            if (fileComparer.AddItems.Count == 0 &&
                fileComparer.MissingItems.Count == 0 &&
                fileComparer.DiffItems.Count == 0)
                return;

            var exportPath = Path.Combine(Directory.GetCurrentDirectory(), ReposrtName);

            if (Directory.Exists(exportPath))
                try
                {
                    Directory.Delete(exportPath, true);
                }
                catch (Exception)
                {
                }

            if (!Directory.Exists(exportPath))
                Directory.CreateDirectory(exportPath);

            var expectedFolder = Path.Combine(exportPath, Expected);
            var outputFolder = Path.Combine(exportPath, Output);

            if (!Directory.Exists(expectedFolder))
                Directory.CreateDirectory(expectedFolder);
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var exportFile = Path.Combine(exportPath, ReposrtName + ".xlsx");
            var epp = new ExcelPackage(new FileInfo(exportFile));
            var sheet = epp.Workbook.Worksheets.Add(ReposrtName);
            sheet.Cells[1, 1].Value = Expected;
            sheet.Cells[1, 2].Value = Output;
            sheet.Cells[1, 3].Value = "Remark";

            var iRow = 2;
            foreach (var item in fileComparer.MissingItems)
            {
                sheet.Cells[iRow, 1].Value = item;
                sheet.Cells[iRow, 3].Value = "Lack of item";
                var tarExpfile = Path.Combine(expectedFolder, Path.GetFileName(item));
                File.Copy(fileComparer.Expected + item, tarExpfile, true);
                iRow++;
            }

            foreach (var item in fileComparer.AddItems)
            {
                sheet.Cells[iRow, 2].Value = item;
                sheet.Cells[iRow, 3].Value = "New item";
                var tarOutfile = Path.Combine(outputFolder, Path.GetFileName(item));
                File.Copy(fileComparer.Output + item, tarOutfile, true);
                iRow++;
            }

            if (fileComparer.DiffItems.Count > 0)
                foreach (var item in fileComparer.DiffItems)
                {
                    var tarExpfinfo = new FileInfo(item.ExpectedFile);
                    var tarExpfile = Path.Combine(expectedFolder, tarExpfinfo.Name);
                    File.Copy(item.ExpectedFile, tarExpfile, true);

                    var tarOutfinfo = new FileInfo(item.OutputFile);
                    var tarOutfile = Path.Combine(outputFolder, tarOutfinfo.Name);
                    File.Copy(item.OutputFile, tarOutfile, true);
                    var bat = Path.Combine(exportPath,
                        "diffpair_" + tarExpfinfo.Name.Replace(tarExpfinfo.Extension, "") + ".bat");

                    sheet.Cells[iRow, 1].Value = tarExpfinfo.Name;
                    sheet.Cells[iRow, 2].Value = tarOutfinfo.Name;
                    sheet.Cells[iRow, 3].Formula = @"=HYPERLINK(""" + bat + @""",""Check the Difference"")";
                    sheet.Cells[iRow, 3].Style.Font.Color.SetColor(Color.Blue);

                    WriteBatFile(bat, "@echo off\nstart " + _kdiff3 + " \"" + tarExpfile + "\" \"" + tarOutfile + "\"");
                    iRow++;
                }

            if (fileComparer.DiffItems.Count > 0)
            {
                var bat2 = Path.Combine(exportPath, "diffpair_all.bat");
                WriteBatFile(bat2, "@echo off\nstart " + _kdiff3 + " " + expectedFolder + " " + outputFolder);
                sheet.Cells[iRow + 2, 3].Formula = @"=HYPERLINK(""" + bat2 + @""",""Check ALL"")";
                sheet.Cells[iRow + 2, 3].Style.Font.Color.SetColor(Color.Blue);
            }

            sheet.Column(1).Width = 40;
            sheet.Column(2).Width = 40;
            sheet.Column(3).Width = 80;
            epp.Save();

            if (File.Exists(exportFile))
            {
                var startInfo = new ProcessStartInfo(exportFile);
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start(startInfo);
            }
        }

        private void WriteBatFile(string outpath, string content)
        {
            var sw = File.CreateText(outpath);
            sw.WriteLine(content);
            sw.Close();
        }

        private bool FileCompare(string file1, string file2)
        {
            int file1Byte;
            int file2Byte;
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
                file1Byte = fs1.ReadByte();
                file2Byte = fs2.ReadByte();
            } while (file1Byte == file2Byte && file1Byte != -1);

            // Close the files.
            fs1.Close();
            fs2.Close();

            // Return the success of the comparison. "file1byte" is
            // equal to "file2byte" at this point only if the files are
            // the same.
            return file1Byte - file2Byte == 0;
        }
    }
}