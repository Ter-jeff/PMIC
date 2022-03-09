using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;

namespace Library.Common
{
    public class Utility
    {
        public static string GetExcelCellValue(ExcelWorksheet sheet, int row, int column)
        {
            if (sheet.Cells[row, column].Formula != null && sheet.Cells[row, column].Formula != "")
                return sheet.Cells[row, column].Formula.ToString();
            if (sheet.Cells[row, column].Value != null)
                return sheet.Cells[row, column].Value.ToString();
            if (sheet.Cells[row, column].Text != null)
                return sheet.Cells[row, column].Text.ToString();
            return "";
        }

        public static void MarkSemiManualCell(ExcelWorksheet sheet, int row, int column)
        {
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }

        public static void MarkManualCell(ExcelWorksheet sheet, int row, int column)
        { 
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);
        }       

        public static string FindInstalledIgxl()
        {
            string igxlRoot = Environment.GetEnvironmentVariable("IGXLROOT");
            FileInfo versionFile = new FileInfo(Path.Combine(igxlRoot, "bin", "Version.txt"));
            if (!versionFile.Exists)
                return string.Empty;

            StreamReader reader = new StreamReader(versionFile.FullName);
            Regex versionReg = new Regex(@"version\s*:\s*(?<Version>[\w\.]+)", RegexOptions.IgnoreCase);
            while (reader.Peek() != -1)
            {
                string line = reader.ReadLine();
                if (versionReg.IsMatch(line))
                {
                    Match versionMatch = versionReg.Match(line);
                    string version = versionMatch.Groups["Version"].Value;
                    reader.Close();
                    reader.Dispose();
                    return version;
                }
            }

            reader.Close();
            reader.Dispose();
            return string.Empty;
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

        public static string ConvertListToString(List<string> inputlist, string joinStr)
        {
            if (inputlist == null || inputlist.Count == 0)
                return "";
            return string.Join(joinStr, inputlist);
        }

        public static bool CompareTwoListItem(List<string> list1, List<string> list2)
        {
            if (list1 == null && list2 == null)
                return true;
            if (list1 == null && list2 != null)
                return false;
            if (list1 != null && list2 == null)
                return false;
            if (list1.Count != list2.Count)
                return false;
            if (list1.Exists(s => !list2.Exists(m => m.Equals(s, StringComparison.OrdinalIgnoreCase))))
                return false;
            if (list2.Exists(s => !list1.Exists(m => m.Equals(s, StringComparison.OrdinalIgnoreCase))))
                return false;
            return true;
        }

        public static DataTable ConvertToDataTable(string filePath, char[] delimiter, string designateToken, int numberOfColumns = 0)
        {
            DataTable table = new DataTable();
            string[] lines = System.IO.File.ReadAllLines(filePath);

            if (numberOfColumns == 0)
            {
                foreach (string line in lines)
                {
                    if (line.Trim().Equals(string.Empty)) continue;
                    if (!designateToken.Equals(string.Empty) && line.IndexOf(designateToken, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        var cols = line.Split(delimiter);
                        numberOfColumns = cols.Count();
                        break;
                    }
                }
            }

            for (int col = 0; col < numberOfColumns; col++)
                table.Columns.Add(new DataColumn("Column" + (col + 1).ToString()));

            foreach (string line in lines)
            {
                var cols = line.Split(delimiter);
                DataRow dr = table.NewRow();               
                for (int cIndex = 0; cIndex < numberOfColumns; cIndex++)
                {
                    if(cIndex < cols.Length)
                        dr[cIndex] = cols[cIndex];
                }

                table.Rows.Add(dr);
            }

            return table;
        }
    }
}
