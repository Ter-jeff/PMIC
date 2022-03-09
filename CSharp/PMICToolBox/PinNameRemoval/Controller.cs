using OfficeOpenXml;
using PinNameRemoval.Operations;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PinNameRemoval
{
    public class Ctrl
    {
        public static List<string> PinsToDelete = new List<string>() { "SPMI_SCLK", "SPMI_SDATA" };
        public static List<int> PinsIndex { get; set; }
        public static List<List<int>> PinIndexList { get; set; }
        public static List<string> tempLog = new List<string>();

        public static List<string> GetAllFilePath(string input, string searchPattern)
        {
            List<string> result = new List<string>();
            FileAttributes attr = File.GetAttributes(input);
            string[] extensions = searchPattern.Split('|');
            if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
            {
                //var files = Directory.EnumerateFiles("C:\\path", "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".mp3") || s.EndsWith(".jpg"));
                foreach (string ext in extensions)
                    result.AddRange(Directory.GetFiles(input, ext).ToList());
                //result.AddRange(Directory.GetFiles(input, searchPattern).ToList());
                List<string> folderList = Directory.GetDirectories(input).ToList();
                foreach (string path in folderList)
                    result.AddRange(GetAllFilePath(path, searchPattern));
            }
            else
                result.Add(input);
            return result;
        }

        public static void SetPinNameListToDelete(string input)
        {
            List<string> pinNames = input.Split(',').ToList();
            PinsToDelete = new List<string>();
            pinNames.ForEach(s => PinsToDelete.Add(s.Trim().ToUpper()));
        }

        public static List<string> GetPinNameList(string input)
        {
            input = input.Substring(input.IndexOf("(") + 1, input.LastIndexOf(")") - input.IndexOf("(") - 1);
            Dictionary<string, string> groupDic = new Dictionary<string, string>();
            int groupIndex = 0;
            while (input.Contains("("))
            {
                string groupItem = input.Substring(input.IndexOf("("), input.IndexOf(")") - input.IndexOf("(") + 1);
                groupDic.Add("GroupItem" + groupIndex, groupItem);
                input = input.Replace(groupItem, "GroupItem" + groupIndex);
                groupIndex++;
            }
            List<string> pinList = input.Replace(" ", "").Split(',').ToList();
            pinList.RemoveAt(0);
            for (int i = 0; i < pinList.Count; i++)
            {
                string s = pinList[i];
                if (groupDic.ContainsKey(s)) pinList[i] = groupDic[s];
            }
            return pinList;
        }

        public static List<List<string>> GetPinList(string input)
        {
            input = input.Substring(input.IndexOf("(") + 1, input.LastIndexOf(")") - input.IndexOf("(") - 1);
            List<List<string>> pinList = new List<List<string>>();
            Dictionary<string, string> groupDic = new Dictionary<string, string>();
            int groupIndex = 0;
            while (input.Contains("("))
            {
                string groupItem = input.Substring(input.IndexOf("("), input.IndexOf(")") - input.IndexOf("(") + 1);
                groupDic.Add("GroupItem" + groupIndex, groupItem.Replace("(", "").Replace(")", ""));
                input = input.Replace(groupItem, "GroupItem" + groupIndex);
                groupIndex++;
            }
            List<string> tempList = input.Replace(" ", "").Split(',').ToList();
            tempList.RemoveAt(0);
            foreach (string s in tempList)
            {
                if (groupDic.ContainsKey(s))
                    pinList.Add(new List<string>(groupDic[s].Replace(" ", "").Split(',').ToList()));
                else
                    pinList.Add(new List<string>() { s });
            }
            return pinList;
        }

        public static void GetPinsIndex(string input)
        {
            List<string> pinNames = GetPinNameList(input);
            PinsIndex = new List<int>();
            Regex pinNamesToDel = new Regex(string.Join("|", PinsToDelete), RegexOptions.IgnoreCase);
            for (int i = 0; i < pinNames.Count; i++)
            {
                if (pinNamesToDel.Match(pinNames[i]).Success)
                    PinsIndex.Add(i);
            }
            PinsIndex.Sort();
            PinsIndex.Reverse();
        }

        public static List<List<int>> GetPinsIndexList(string input)
        {
            List<List<string>> pinNames = GetPinList(input);
            List<List<int>> pinIndexList = new List<List<int>>();
            List<string> tempLst = new List<string>();
            foreach(var x in PinsToDelete)
            {
                tempLst.Add("^" + x + "$");
            }
            Regex pinNamesToDel = new Regex(string.Join("|", tempLst), RegexOptions.IgnoreCase);
            foreach (List<string> obj in pinNames)
            {
                List<int> temp = new List<int>();
                for (int i = 0; i < obj.Count; i++)
                {
                    if (pinNamesToDel.Match(obj[i]).Success)
                        temp.Add(i);
                }
                pinIndexList.Add(temp);
            }
            return pinIndexList;
        }

        public List<string> DeletePins(string filePath, out bool isScan)
        {
            List<string> resultLines = new List<string>();
            string line;
            isScan = false;
            List<string> logStrings = new List<string>();
            logStrings.Add("----------------------------------------------------------");
            logStrings.Add(filePath);
            int index = 0;
            using (StreamReader sr = new StreamReader(filePath))
            {
                Operation oper;
                while ((line = sr.ReadLine()) != null)
                {
                    oper = Operation.CreateOperation(line);
                    if (oper.GetType() != typeof(Operation))
                        logStrings.Add(string.Format("  line {0}: {1}", index, oper.GetType().ToString()));
                    if (oper.GetType() == typeof(ScanPinsOperation))
                        isScan = true;
                    resultLines.AddRange(oper.RemovePins(ref index, sr));
                }
            }
            tempLog.AddRange(logStrings);
            return resultLines;
        }

        public static bool CompilerCLI(string inputPath, string outputPath, string switches, out string outputMsg)
        {
            outputMsg = string.Empty;
            int exitCode = -1;
            tempLog.Add("CLI-----------------------");
            tempLog.Add("CLI  inputPath: " + inputPath);
            tempLog.Add("CLI  outputPath: " + outputPath);
            tempLog.Add("CLI  switches: " + switches);
            using (Process p = new Process())
            {
                p.StartInfo = new ProcessStartInfo("cmd.exe")
                {
                    Arguments = switches,
                    //WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                };
                p.Start();
                p.WaitForExit();
                outputMsg = p.StandardOutput.ReadToEnd();
                exitCode = p.ExitCode;
            }

            if (exitCode == 0) return true;
            else return false;
        }

        public static List<ReverseCompile> CompileAsync(List<ReverseCompile> compileList)
        {
            List<Task> tpl = new List<Task>();
            foreach (var item in compileList)
            {
                Task t = Task.Run(() => item.Run());
                tpl.Add(t);
            }
            Task.WaitAll(tpl.ToArray());
            foreach (var res in compileList)
                tempLog.Add(string.Format(res.Result + "\t" + res.FileName + "\t" + res.OutputMsg));
            return compileList;
        }

        public string Process(PatternInfo pi)
        {
            pi.APRC.Run();
            bool isScan;
            List<string> deleteResult = new List<string>();
            try
            {
                deleteResult = DeletePins(pi.APRC.OutputPath, out isScan);
                pi.APC.isScanType = isScan;
            }
            catch (Exception ex)
            {
                pi.APRC.OutputMsg = ex.Message;
            }
            string targetPath = pi.APRC.OutputPath.Replace(pi.FileNameWithoutExt, pi.FileNameWithoutExt + "_beforeRemoval");
            File.Copy(pi.APRC.OutputPath, targetPath, true);
            File.Delete(pi.APRC.OutputPath);
            File.WriteAllLines(pi.APRC.OutputPath, deleteResult);
//#if !DEBUG
            pi.APC.Run();
//#endif
            return string.Format("{0}\t{1}\t{2}", pi.APRC.FileName, pi.APRC.Result.ToString() + " " + pi.APRC.OutputMsg, pi.APC.Result.ToString() + " " + pi.APC.OutputMsg);
        }

        public async Task<List<string>> ProcessAsync(List<PatternInfo> patternList, IProgress<int> progress)
        {
            List<string> msgs = new List<string>();
            //List<Task> tpl = new List<Task>();
            int count = 0;

            foreach (var item in patternList)
            {
                await Task.Run(() =>
                  {
                      msgs.Add(Process(item));
                      if (progress != null)
                          progress.Report(count + 1);
                      count++;
                  });
                //tpl.Add(t);
            }
            //await Task.WhenAll(tpl);
            return msgs;
        }

        static string SucceedOrFailed(bool input)
        {
            return input ? "Succeed" : "Failed";
        }

        public void GenerateReport(List<PatternInfo> patternList)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add("Wroksheet1");
                List<string[]> rows = new List<string[]>() { new string[] { "Item", "File Name", "Scan Type", "Result", "Path", "Reverse Compile Result", "Output Message", "Compile Result", "Output Message" } };
                int index = 1;
                foreach (PatternInfo pi in patternList)
                {
                    string num = index++.ToString();
                    string isScan = pi.APC.isScanType ? "Yes" : string.Empty;
                    string result = SucceedOrFailed(pi.APRC.Result && pi.APC.Result);
                    string aprcResult = SucceedOrFailed(pi.APRC.Result);
                    string apcResult = SucceedOrFailed(pi.APC.Result);
                    rows.Add(new string[] { num, pi.APRC.FileName, isScan, result, pi.InputPath, aprcResult, pi.APRC.OutputMsg, apcResult, pi.APC.OutputMsg });
                }

                string range = "A1:" + char.ConvertFromUtf32(rows[0].Length + 64) + rows.Count;
                worksheet.Cells[range].LoadFromArrays(rows);
                //worksheet.Cells[range].AutoFitColumns();
                worksheet.Column(2).Width = 30;
                worksheet.Column(5).Width = 50;
                worksheet.Column(6).Width = 20;
                worksheet.Column(7).Width = 20;
                worksheet.Column(8).Width = 15;
                worksheet.Column(9).Width = 20;

                FileInfo excelFile = new FileInfo(Path.Combine(Path.GetDirectoryName(patternList[0].OutputPath), "PinRemovalResult_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx"));
                excel.SaveAs(excelFile);
            }
        }
    }
}
