using AutomationCommon.DataStructure;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Teradyne.Oasis.IGLinkBase;
using Teradyne.Oasis;

namespace IgxlData.IgxlManager
{
    public class IgxlManagerMain
    {
        public ManifestSheet Manifest;
        private const string Iglink = "IGLink";

        public bool ExportWorkBook(string testProgramName, string exportFolder)
        {
            if (!Directory.Exists(exportFolder))
                Directory.CreateDirectory(exportFolder);
            else
            {
                Directory.Delete(exportFolder, true);
                Directory.CreateDirectory(exportFolder);
            }
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            string exportWorkBookCmd = oasisRootFolder + @"ExportWorkbook.exe";
            if (File.Exists(exportWorkBookCmd))
            {
                string option = "-w \"" + testProgramName + "\" -d \"" + exportFolder + "\"";
                return RunCmd(exportWorkBookCmd, option);
            }
            return false;
        }

        public string GenerateIgxlProgram(string exportFolder, string outputPath, double version)
        {
            //Step1 Get LinkStructure
            CreateIgLinkStructure(outputPath, exportFolder, version);
            //Step2 Generate by project file
            return GenIgxlProgramByProjectFile(outputPath, version);
        }

        public void GenIgxlProgram(List<string> sourceFiles, string outputFolder, string subProgramName, IgxlWorkBook igxlWorkBook, Action<string, MessageLevel, int> report, string version)
        {
            try
            {
                DeviceProject nProject = new DeviceProject();
                SubProgram nSubProgram = new SubProgram();

                string tmpIgLinkFolder = Path.Combine(outputFolder, @"IGLink");
                if (!Directory.Exists(tmpIgLinkFolder)) Directory.CreateDirectory(tmpIgLinkFolder);

                PropertyInfo fileNameInfo= nProject.GetType().GetProperty("FileName");
                fileNameInfo.SetValue(nProject, tmpIgLinkFolder + @"\" + subProgramName + @".igxlProj", null);
                nProject.Name = "ProjectTemple";
                //nProject.FileName = tmpIgLinkFolder + @"\" + subProgramName + @".igxlProj";
                nProject.CurrentDir = tmpIgLinkFolder;
                nProject.SaveAsXLS = false;
                nProject.SheetOrder = SheetOrderPreference.Alphabetically;
                
                nSubProgram.Name = subProgramName;
                nSubProgram.JobNames = subProgramName;

                foreach (string sourceFile in sourceFiles)
                {
                    if (sourceFile.ToLower().IndexOf("VBT_Instrument_Setup", StringComparison.OrdinalIgnoreCase) != -1)
                        continue;
                    if (sourceFile.ToLower().IndexOf(".txt", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        nSubProgram.Add(GetSheetInfo(sourceFile, tmpIgLinkFolder));
                    }
                    else if (sourceFile.ToLower().IndexOf(".bas", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        nSubProgram.Add(GetVbInfo(sourceFile, tmpIgLinkFolder));
                    }
                    else if (sourceFile.ToLower().IndexOf(".cls", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        nSubProgram.Add(GetVbInfo(sourceFile, tmpIgLinkFolder));
                    }
                    else if (sourceFile.ToLower().IndexOf(".frm", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        nSubProgram.Add(GetVbInfo(sourceFile, tmpIgLinkFolder));
                    }
                    else if (sourceFile.ToLower().IndexOf(".frx", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        nSubProgram.Add(GetVbInfo(sourceFile, tmpIgLinkFolder));
                    }
                }

                string defaultFlow = "";
                foreach (KeyValuePair<string, FlowSheet> pair in igxlWorkBook.MainFlowSheets)
                {
                    if (defaultFlow == "") defaultFlow = pair.Value.SheetName;
                }
                nSubProgram.MainFlow = subProgramName + ":" + defaultFlow;
                nSubProgram.GenerateJobListSheet = false;
                nSubProgram.TargetIGXLVersion = GetTargetIgxlVersion(version);
                nSubProgram.ExportIGXLVersion = GetTargetIgxlVersion(version);
                nSubProgram.FlowEntry= subProgramName + ":" + defaultFlow;

                nProject.SubPrograms.Add(nSubProgram);

                report("Saving IGLink Project ...", MessageLevel.General, 40);

                DeviceProject.SaveProjectCfg(nProject);

                double versionDouble = double.Parse(version.Substring(1));
                if (versionDouble < 9.0)
                {
                    report("Generating IGXL workbook ...", MessageLevel.General, 80);
                    GenerateIgxlProgram(Path.Combine(tmpIgLinkFolder, subProgramName + @".igxlProj"), Path.Combine(tmpIgLinkFolder, nSubProgram.Name + ".xlsm"), nSubProgram.Name, " -e ", "");
                }
                else
                {
                    report("Creating IGXL workbook ...", MessageLevel.General, 80);
                    GenerateIgxlProgram(Path.Combine(tmpIgLinkFolder, subProgramName + @".igxlProj"), Path.Combine(tmpIgLinkFolder, nSubProgram.Name + ".igxl"), nSubProgram.Name, " -g ", "");
                    GenerateIgxlProgram(Path.Combine(tmpIgLinkFolder, subProgramName + @".igxlProj"), Path.Combine(tmpIgLinkFolder, nSubProgram.Name + ".xlsm"), nSubProgram.Name, " -e ", "");
                }

            }
            catch (Exception e)
            {
                report("Error occurs during generate IGLink Project" + e.Message, MessageLevel.Error, 100);
            }
        }

        public void CreateIgLinkStructure(string outputPath, string exportFolder, double version)
        {
            Manifest = Manifest ?? ReadManifestSheet(exportFolder, Iglink, outputPath);
            var unknownIglinkFilePath = Directory.GetFiles(exportFolder).ToList();
            foreach (string file in Directory.GetFiles(exportFolder))
            {
                string fileName = Path.GetFileName(file);
                var ext = Path.GetExtension(fileName);
                fileName = fileName.Replace(ext, ext.ToLower());
                if (!Manifest.Items.Exists(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)))
                    UpdateFileStructureWithUnknownFiles(fileName, Manifest, Path.Combine(outputPath, Iglink));

                if (Manifest.Items.Exists(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var item = Manifest.Items.Find(x => x.FileName.Equals(fileName, StringComparison.CurrentCultureIgnoreCase));
                    string newFileName = Path.Combine(outputPath, item.FullFilePath);
                    string folder = Path.GetDirectoryName(newFileName);
                    if (folder != null && !Directory.Exists(folder))
                        Directory.CreateDirectory(folder);
                    File.Copy(file, newFileName, true);
                    unknownIglinkFilePath.Remove(file);
                }
            }

            string mainFlow = GetFlowMain90(outputPath);

            if (!string.IsNullOrEmpty(mainFlow))
                Manifest.MainFlow = mainFlow;

            Manifest.ArrangeSubprogramItems();
        }

        private static string GetFlowMain90(string outputFolder)
        {
            var manifestFile = Path.Combine(outputFolder + @"\exportProg", "JobList.txt");
            if (!File.Exists(manifestFile)) return "";
            var allFileInfo = File.ReadAllLines(manifestFile).ToList();
            int flowTableIndexCol = -1;
            int flowTableIndexRow = -1;
            for (int i = 3; i < allFileInfo.Count; i++)
            {
                var line = allFileInfo[i].Split('\t').ToList();
                for (var j = 0; j < line.Count; j++)
                {
                    if (line[j].Equals("Flow Table", StringComparison.OrdinalIgnoreCase))
                    {
                        flowTableIndexRow = i;
                        flowTableIndexCol = j;
                        break;
                    }
                }
            }
            if (flowTableIndexRow == -1) return "";
            for (var i = flowTableIndexRow + 1; i < allFileInfo.Count; i++)
            {
                var line = allFileInfo[i].Split('\t').ToList();
                if (line[flowTableIndexCol] == null)
                    continue;
                return line[flowTableIndexCol];
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

            var trunkPath = manifest.Items.Count == 0 ? Iglink :
                Path.Combine(targetFolder, manifest.Items.First().RelatedPath.Split('/')[0]);

            var newPath = Path.Combine(trunkPath, "Unknown", filename);
            manifest.UpdateManifestItemWithUnknownItem("Unknown", filename, Path.Combine(trunkPath.Split('\\').ToList().Last(), "Unknown", filename), newPath);
        }

        private static ManifestSheet ReadManifestSheet(string outputFolder, string iglink, string targetFolder)
        {
            var manifest = new ManifestSheet(targetFolder);
            manifest.TargetFolder = targetFolder;
            var manifestFile = Path.Combine(outputFolder, "IGLinkManifest.txt");
            if (!File.Exists(manifestFile)) return manifest;
            var allFileInfo = File.ReadAllLines(manifestFile).ToList();
            var flagCollectItems = false;

            for (var i = 0; i <= allFileInfo.Count - 1; i++)
            {
                var line = allFileInfo[i].Split('\t').ToList();
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

        public string GenIgxlProgramByProjectFile(string outFolder, double version)
        {
            var outSubFolder = outFolder + @"\" + Iglink;
            Directory.CreateDirectory(outSubFolder);

            //STEP1. Set DeviceProject
            SubProgram objSubProgram = new SubProgram();
            var igLinkProjectPath = outSubFolder + @"\" + Manifest.ProjectName + @".igxlProj";

            var objProject = new DeviceProject();
            //objProject.JobNames = Manifest.JobName;
            objProject.Name = Manifest.ProjectName;
            //objProject.FileName = igLinkProjectPath;
            objProject.SaveAsXLS = true;
            objProject.SheetOrder = SheetOrderPreference.Alphabetically;
            objProject.SaveAsXLS = false;

            //STEP2. Set SubProgram
            objSubProgram.Name = Manifest.ProjectName;
            objSubProgram.JobNames = Manifest.JobName;
            var newFiles = new List<string>();
            foreach (var subprogram in Manifest.SubPrograms)
            {
                foreach (var file in subprogram.Value)
                {
                    if (!newFiles.Contains(file.RelatedPath))
                        newFiles.Add(file.RelatedPath);
                }
            }
            //2.1 Search all files and add into subProgram
            foreach (var srcFile in newFiles)
            {
                if (srcFile.ToLower().IndexOf(".txt", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetSheetInfo(srcFile, ""));

                else if (srcFile.ToLower().IndexOf(".bas", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetVbInfo(srcFile, ""));

                else if (srcFile.ToLower().IndexOf(".cls", StringComparison.Ordinal) != -1)
                    objSubProgram.Add(GetVbInfo(srcFile, ""));
            }

            //2.2 Set SubProgram.MainFlow
            objSubProgram.MainFlow = Manifest.ProjectName + ":" + Manifest.MainFlow;
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

            var igxlProgramPath = outSubFolder + @"\" + Manifest.ProjectName + ext;
            //3.1 Generate IG-Link project and iG-Excel program
            GenerateIgxlProgram(igLinkProjectPath, igxlProgramPath, objSubProgram.Name, switchFlag, Manifest.JobName);
            if (version != 8.3)
                GenerateIgxlProgram(igLinkProjectPath, igxlProgramPath, Path.ChangeExtension(objSubProgram.Name, ".xlsm"), switchFlag, Manifest.JobName);
            return File.Exists(igxlProgramPath) ? igxlProgramPath : "";
        }

        private Sheet GetSheetInfo(string filepath, string refPath)
        {
            Sheet nSheet = new Sheet();
            string tmpStr;
            if (refPath != "")
            {
                tmpStr = filepath.ToLower().IndexOf(refPath.ToLower(), StringComparison.Ordinal) != -1 ?
                    filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1) :
                    filepath;
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
            string tmpStr;
            if (refPath != "")
            {
                tmpStr = filepath.ToLower().IndexOf(refPath.ToLower(), StringComparison.Ordinal) != -1 ? 
                    filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1) :
                    filepath;
            }
            else
            {
                tmpStr = filepath;
            }
            nVbModule.Source = tmpStr;
            return nVbModule;
        }

        private void GenerateIgxlProgram(string linkProject, string outputFile, string subProgramName, string switchStr, string jobName)
        {
            string option = "";
            var oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            var igLinkCl = oasisRootFolder + @"IGLinkCL.exe";
            if (File.Exists(igLinkCl))
            {
                if (subProgramName != "")
                    option = "-i " + "\"" + linkProject + "\"" + " -s " + "\"" + subProgramName + "\"" + switchStr + "\"" + outputFile + "\"";
                else if (jobName != "") option = "-i " + "\"" + linkProject + "\"" + " -j " + "\"" + subProgramName + "\"" + switchStr + "\"" + outputFile + "\"";
                GenerateIgxl(igLinkCl, option);
            }
        }

        private void GenerateIgxl(string cmd, string argument)
        {
            Process nProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Normal;
            startInfo.FileName = cmd;
            startInfo.Arguments = argument;
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        private bool RunCmd(string cmd, string argument = "")
        {
            Process nProcess = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.FileName = cmd;
            startInfo.Arguments = argument;
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
            if (nProcess.ExitCode == 0)
                return true;
            return false;
        }

        private SupportableIGXLVersions GetTargetIgxlVersion(string version)
        {
            if (version.StartsWith("v8.30",StringComparison.OrdinalIgnoreCase))
                return SupportableIGXLVersions.V8_30_ultraflex;
            else if (version.StartsWith("v9.00",StringComparison.OrdinalIgnoreCase))
                return SupportableIGXLVersions.V9_00_ultraflex;
            else if (version.StartsWith("v10.00",StringComparison.OrdinalIgnoreCase))
                return SupportableIGXLVersions.V10_00_ultraflex;
            else if (version.StartsWith("v10.20",StringComparison.OrdinalIgnoreCase))
                return SupportableIGXLVersions.V10_20_ultraflex;
            else
                return SupportableIGXLVersions.Default;
        }
    }
}