using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Teradyne.Oasis.IGLinkBase;

namespace IgxlData.IgxlManager
{
    public class IgxlManagerMain
    {
        private const string Iglink = "IGLink";
        public ManifestSheet Manifest { get; set; }

        public bool ExportWorkBook(string testProgramFile, string exportfolder)
        {
            if (!Directory.Exists(exportfolder))
                Directory.CreateDirectory(exportfolder);
            else
            {
                Directory.Delete(exportfolder, true);
                Directory.CreateDirectory(exportfolder);
            }

            //if (Path.GetExtension(testProgramFile).Equals(".igxl", StringComparison.CurrentCultureIgnoreCase))
            //{
            //    ZipFile.ExtractToDirectory(testProgramFile, exportfolder);
            //}
            //else
            {
                string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
                string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
                if (File.Exists(exportWorkBookCmd))
                {
                    string option = "-w \"" + testProgramFile + "\" -d \"" + exportfolder + "\"";
                    return RunCmd(exportWorkBookCmd, option);
                }
            }
            return false;
        }

        public void GenXlsxByTxt(string testprogramName, List<string> files)
        {
            files = files.OrderBy(Path.GetFileName).ToList();
            if (File.Exists(testprogramName))
                File.Delete(testprogramName);
            using (var excel = new ExcelPackage(new FileInfo(testprogramName)))
            {
                foreach (var file in files)
                {
                    var worksheetsName = Path.GetFileNameWithoutExtension(file);
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(worksheetsName);
                    var format = new ExcelTextFormat();
                    format.Delimiter = '\t';
                    worksheet.Cells[1, 1].LoadFromText(new FileInfo(file), format);
                }
                excel.Save();
            }
        }

        public void GenTestProgramByTxt(string outFolder, string testProgramFile, double version)
        {
            var ext = @".xlsm";
            if (!(version < 9.0))
                ext = @".igxl";

            if (testProgramFile != null)
            {
                testProgramFile = Path.ChangeExtension(testProgramFile, ext);
                string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
                string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
                if (File.Exists(exportWorkBookCmd))
                {
                    string option = "-w \"" + outFolder + "\" -i \"" + testProgramFile + "\"";
                    RunCmd(exportWorkBookCmd, option);
                }
            }
        }

        public string GenTestProgramByManifest(string outputPath, string exportFolder, double version)
        {
            //Step1 Get Manifest
            Manifest = ReadManifestSheet(exportFolder, Iglink, outputPath);
            //Step2 Get LinkStructure
            CreateIgLinkStructure(outputPath, exportFolder, Manifest);
            //Step3 Generate by project file
            return GenIgxlProgramByProjectFile(outputPath, version, Manifest);
        }

        public void GenTestProgramByProject(string glinkProject, string outputFile, string subProgramName, string switchStr, string jobName)
        {
            string option = "";
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            string glinkCl = oasisRootFolder + @"IGLinkCL.exe";
            if (File.Exists(glinkCl))
            {
                if (subProgramName != "")
                    option = "-i " + "\"" + glinkProject + "\"" + " -s " + "\"" + subProgramName + "\"" + switchStr + "\"" + outputFile + "\"";
                else if (jobName != "")
                    option = "-i " + "\"" + glinkProject + "\"" + " -j " + "\"" + subProgramName + "\"" + switchStr + "\"" + outputFile + "\"";
                GenerateIgxl(glinkCl, option);
            }
        }

        public string ConvertIgxlToExcel(string testProgram, string exportExceltestProgram)
        {
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
            if (File.Exists(exportWorkBookCmd))
            {
                string option = "-w \"" + testProgram + "\" -e \"" + exportExceltestProgram + "\"";
                RunCmd(exportWorkBookCmd, option);
            }
            return exportExceltestProgram;
        }

        public double GetVersion(string fileName)
        {
            return Path.GetExtension(fileName).ToLower() == ".igxl" ? 9.0 : 8.3;
        }

        private void CreateIgLinkStructure(string outputPath, string exportFolder, ManifestSheet manifest)
        {
            var unknownIGlinkFilePath = Directory.GetFiles(exportFolder).ToList();
            foreach (string file in Directory.GetFiles(exportFolder))
            {
                string fileName = Path.GetFileName(file);
                var ext = Path.GetExtension(fileName);
                fileName = fileName.Replace(ext, ext.ToLower());
                if (fileName != null)
                {
                    if (!manifest.Items.Exists(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)))
                        UpdateFileStructureWithUnknownFiles(fileName, manifest, Path.Combine(outputPath, Iglink));

                    if (manifest.Items.Exists(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var item = manifest.Items.Find(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase));
                        string newFileName = Path.Combine(outputPath, item.FullFilePath);
                        string folder = Path.GetDirectoryName(newFileName);
                        if (!Directory.Exists(folder))
                            Directory.CreateDirectory(folder);
                        File.Copy(file, newFileName, true);
                        unknownIGlinkFilePath.Remove(file);
                    }
                }
            }

            //string mainFlow = version < 9.0 ? GetFlowMain(progPath) : GetFlowMain90(targetFolder);
            string mainFlow = GetFlowMain90(outputPath);

            if (!string.IsNullOrEmpty(mainFlow))
                manifest.MainFlow = mainFlow;

            manifest.ArrangeSubprogramItems();
        }

        private string GetFlowMain90(string outputFolder)
        {
            var manifestfile = Path.Combine(outputFolder + @"\exportProg", "JobList.txt");
            if (!File.Exists(manifestfile)) return "";
            var allfileInfo = File.ReadAllLines(manifestfile).ToList();
            int flowTalbeIndexCol = -1;
            int flowTalbeIndexRow = -1;
            for (int i = 3; i < allfileInfo.Count; i++)
            {
                var line = allfileInfo[i].Split('\t').ToList();
                for (var j = 0; j < line.Count(); j++)
                {
                    if (line[j].Equals("Flow Table", StringComparison.OrdinalIgnoreCase))
                    {
                        flowTalbeIndexRow = i;
                        flowTalbeIndexCol = j;
                        break;
                    }
                }
            }
            if (flowTalbeIndexRow == -1) return "";
            for (var i = flowTalbeIndexRow + 1; i < allfileInfo.Count; i++)
            {
                var line = allfileInfo[i].Split('\t').ToList();
                if (line[flowTalbeIndexCol] == null)
                    continue;
                return line[flowTalbeIndexCol];
            }
            var name = Directory.GetFiles(outputFolder, "*.txt", SearchOption.TopDirectoryOnly).ToList().Find(x => x.StartsWith("Main_Flow_", StringComparison.CurrentCultureIgnoreCase));
            return name;
        }

        private void UpdateFileStructureWithUnknownFiles(string filename, ManifestSheet manifest, string targetFolder)
        {
            if (Regex.IsMatch(filename, "IGLinkManifest", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(filename, "^_", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(filename, "workbook", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(filename, ".bas|.cls", RegexOptions.IgnoreCase))
                return;
            if (manifest.Items.Count == 0)
                return;
            var trunkpath = Path.Combine(targetFolder, manifest.Items.First().RelatedPath.Split('/')[0]); // filestructure.ToList()[0].Split('/')[0];//Arbitrary access 1 path to search trunk folder
            var newpath = Path.Combine(trunkpath, "Unknown", filename);
            manifest.UpdateManifestItemWithUnknownItem("Unknown", filename, Path.Combine(trunkpath.Split('\\').ToList().Last(), "Unknown", filename), newpath);
            //filestructure.Add(newpath);
        }

        private static ManifestSheet ReadManifestSheet(string outputFolder, string iglink, string targetFolder)
        {
            var manifest = new ManifestSheet(targetFolder);
            manifest.TargetFolder = targetFolder;
            var manifestfile = Path.Combine(outputFolder, "IGLinkManifest.txt");
            if (!File.Exists(manifestfile)) return manifest;
            var allfileInfo = File.ReadAllLines(manifestfile).ToList();
            var flagCollectItems = false;

            for (var i = 0; i <= allfileInfo.Count - 1; i++)
            {
                var line = allfileInfo[i].Split('\t').ToList();
                if (line.Count == 1) continue;
                if (line[0] == "IG-Link Project File:")
                {
                    manifest.UpdateManifestHeaderInfo(line[1]);
                    continue;
                }

                if (line[0].StartsWith("Generated by IG-Link", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!flagCollectItems)
                    flagCollectItems = manifest.IsValidManifestItem(line);
                else
                {
                    if (line.Count >= 3)
                        manifest.UpdateManifestItem(line, iglink);
                }
            }
            return manifest;
        }

        private string GenIgxlProgramByProjectFile(string outFolder, double version, ManifestSheet manifest)
        {
            var outSubFolder = outFolder + @"\" + Iglink;
            Directory.CreateDirectory(outSubFolder);

            //STEP1. Set DeviceProgject
            SubProgram objSubProgram = new SubProgram();
            var igLinkProjectPath = outSubFolder + @"\" + manifest.ProjectName + @".igxlProj";

            var objProject = new DeviceProject();
            objProject.JobNames = manifest.JobName;
            objProject.Name = manifest.ProjectName;
            objProject.FileName = igLinkProjectPath;
            objProject.SaveAsXLS = true;
            objProject.SheetOrder = SheetOrderPreference.Alphabetically;
            objProject.SaveAsXLS = false;

            //STEP2. Set SubProgram
            objSubProgram.Name = manifest.ProjectName;
            objSubProgram.JobNames = manifest.JobName;
            var newfiles = new List<string>();
            foreach (var subprogram in manifest.SubPrograms)
            {
                foreach (var file in subprogram.Value)
                {
                    if (!newfiles.Contains(file.RelatedPath))
                        newfiles.Add(file.RelatedPath);
                }
            }
            //2.1 Search all files and add into subProgram
            foreach (var srcFile in newfiles)
            {
                if (srcFile.ToLower().IndexOf(".txt", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetSheetInfo(srcFile, ""));

                else if (srcFile.ToLower().IndexOf(".bas", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetVbInfo(srcFile, ""));

                else if (srcFile.ToLower().IndexOf(".cls", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetVbInfo(srcFile, ""));
            }

            //2.2 Set SubProgram.MainFlow
            objSubProgram.MainFlow = manifest.ProjectName + ":" + manifest.MainFlow;
            objSubProgram.GenerateJobListSheet = false;
            objProject.SubPrograms.Add(objSubProgram);

            DeviceProject.SaveProjectCfg(objProject);

            var switchFlag = " -e ";
            var ext = @".xlsm";
            if (!(version < 9.0))
            {
                switchFlag = " -g ";
                ext = @".igxl";
            }

            var igexProramPath = outSubFolder + @"\" + manifest.ProjectName + ext;
            //3.1 Generate IG-Link project and iG-Excel program
            GenTestProgramByProject(igLinkProjectPath, igexProramPath, objSubProgram.Name, switchFlag, manifest.JobName);

            return File.Exists(igexProramPath) ? igexProramPath : "";
        }

        private Sheet GetSheetInfo(string filepath, string refPath)
        {
            Sheet nSheet = new Sheet();
            string tmpStr = "";
            if (refPath != "")
            {
                if (filepath.ToLower().IndexOf(refPath.ToLower()) != -1)
                {
                    tmpStr = filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1);
                }
                else
                {
                    tmpStr = filepath;
                }
            }
            else
            {
                tmpStr = filepath;
            }
            nSheet.Source = tmpStr;
            return nSheet;
        }

        private VBFile GetVbInfo(string filepath, string refPath)
        {
            VBFile nVbModule = new VBFile();
            string tmpStr = "";
            if (refPath != "")
            {
                if (filepath.ToLower().IndexOf(refPath.ToLower(), StringComparison.Ordinal) != -1)
                {
                    tmpStr = filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1);
                }
                else
                {
                    tmpStr = filepath;
                }
            }
            else
            {
                tmpStr = filepath;
            }
            nVbModule.Source = tmpStr;
            return nVbModule;
        }
       
        private bool GenerateIgxl(string cmdstr, string argment)
        {
            Process nProcess = new Process();
            ProcessStartInfo StartInfo = new ProcessStartInfo();
            StartInfo.UseShellExecute = false;
            StartInfo.RedirectStandardOutput = true;
            StartInfo.CreateNoWindow = true;
            StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            StartInfo.FileName = cmdstr;
            StartInfo.Arguments = argment;
            nProcess.StartInfo = StartInfo;
            nProcess.Start();
            nProcess.WaitForExit();
            if (nProcess.ExitCode == 0)
                return true;
            return false;
        }

        private bool RunCmd(string cmd, string argment = "")
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
            if (nProcess.ExitCode == 0)
                return true;
            return false;
        }
    }
}