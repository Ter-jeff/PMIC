using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using PmicAutogen.InputPackages.Base;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Local;

namespace PmicAutogen.InputPackages
{
    public abstract class InputPackageBase
    {
        public delegate void WriteMessage(string msg, MessageLevel level = MessageLevel.General, int percentage = -1);

        public string Project = "";

        protected InputPackageBase()
        {
            InputFiles = new List<Input>();
        }

        public List<Input> InputFiles { get; set; }

        protected List<InputTestPlan> SelectedTestPlan
        {
            get
            {
                var selected = InputFiles.FindAll(p => p is InputTestPlan && p.Selected);
                return selected.Cast<InputTestPlan>().ToList();
            }
            set { throw new NotImplementedException(); }
        }

        protected List<InputVbtGenTool> SelectedVbtTestPlan
        {
            get
            {
                var selected = InputFiles.FindAll(p => p is InputVbtGenTool && p.Selected);
                return selected.Cast<InputVbtGenTool>().ToList();
            }
            set { throw new NotImplementedException(); }
        }

        protected List<InputScgh> SelectedScgh
        {
            get
            {
                var selected = InputFiles.FindAll(p => p is InputScgh && p.Selected);
                return selected.Cast<InputScgh>().ToList();
            }
            set { throw new NotImplementedException(); }
        }

        protected List<InputPatternListCsv> SelectedPatternList
        {
            get
            {
                var selected = InputFiles.FindAll(p => p is InputPatternListCsv && p.Selected);
                return selected.Cast<InputPatternListCsv>().ToList();
            }
            set { throw new NotImplementedException(); }
        }

        protected List<string> SelectedOtpRegisterMap
        {
            get
            {
                return InputFiles.Where(p => p.FileType == InputFileType.OtpRegisterMap).Select(p => p.FullName)
                    .ToList();
            }
            set { throw new NotImplementedException(); }
        }

        public void ReadFiles(List<string> files, IProgress<ProgressStatus> progress = null)
        {
            var duplicateInputFile = files.Find(file =>
                InputFiles.Exists(x => x.FullName.Equals(file, StringComparison.OrdinalIgnoreCase)));

            if (duplicateInputFile != null)
                throw new Exception("[Duplicated] file already existed! " + duplicateInputFile);

            Project = GetCurrentProjectName(files);
            LocalSpecs.CurrentProject = Project;

            var inputs = new List<Input>();
            foreach (var file in files)
            {
                var info = new FileInfo(file);
                var inputFile = ReadOneFile(info);
                if (inputFile != null)
                {
                    inputs.Add(inputFile);
                    if (inputFile is ExcelInput)
                    {
                        var excelInput = (ExcelInput) inputFile;
                        excelInput.AnalyzeInput();
                    }
                }
                else
                {
                    throw new Exception("Unknown File Type: " + file);
                }
            }

            InputFiles.AddRange(inputs);
        }

        private Input ReadOneFile(FileInfo fileInfo)
        {
            var inputFileType = GetFileType(fileInfo.FullName);
            switch (inputFileType)
            {
                case InputFileType.TestPlan:
                    return new InputTestPlan(fileInfo);
                case InputFileType.VbtGenTool:
                    return new InputVbtGenTool(fileInfo);
                case InputFileType.ScghPatternList:
                    return new InputScgh(fileInfo);
                case InputFileType.PatternListCsv:
                    return new InputPatternListCsv(fileInfo);
                case InputFileType.OtpRegisterMap:
                    return new InputOtpRegisterMap(fileInfo);
                default:
                    return null;
            }
        }

        private InputFileType GetFileType(string fileName)
        {
            var fileFormat = InputFileType.Unknown;
            var fInfo = new FileInfo(fileName);

            switch (fInfo.Extension.ToLower()) //從副檔名判斷
            {
                case ".xlsm":
                    if (Regex.IsMatch(fInfo.Name, @"VBTPOP_Gen_", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.VbtGenTool;
                    else if (Regex.IsMatch(fInfo.Name, @"_Test.*Plan", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.TestPlan;
                    else if (Regex.IsMatch(fInfo.Name, @"_scgh", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.ScghPatternList;
                    else if (Regex.IsMatch(fInfo.Name, @"Bin_Cut", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.BinCut;
                    else fileFormat = InputFileType.IgxlTestProgram;
                    break;
                case ".igxl":
                    fileFormat = InputFileType.IgxlTestProgram;
                    break;
                case ".xlsx":
                    if (Regex.IsMatch(fInfo.Name, @"_scgh", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.ScghPatternList;
                    else if (Regex.IsMatch(fInfo.Name, @"_pat_scg", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.ScghPatternList;
                    else if (Regex.IsMatch(fInfo.Name, @"Bin_Cut", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.BinCut;
                    else if (Regex.IsMatch(fInfo.Name, @"_Test.*Plan", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.TestPlan;
                    else if (Regex.IsMatch(fInfo.Name, @"eFuse_BitDef_", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.EFuseBitDefinition;
                    else if (Regex.IsMatch(fInfo.Name, @"ExceptionList", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.ExceptionList;
                    else if (Regex.IsMatch(fInfo.Name, @"HardIP_TTR", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.HardIpTtr;
                    else if (Regex.IsMatch(fInfo.Name, @"VBTPOP_Gen_", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.VbtGenTool;
                    break;
                case ".csv":
                    if (Regex.IsMatch(fInfo.Name, @"Patterns*", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.PatternListCsv;
                    if (Regex.IsMatch(fInfo.Name, @"PatternList", RegexOptions.IgnoreCase))
                        fileFormat = InputFileType.PatternListCsv;
                    break;
                case ".otp":
                case ".yaml":
                    fileFormat = InputFileType.OtpRegisterMap; //檢查文字檔內容
                    break;
                case ".gz":
                    if (Regex.IsMatch(fInfo.Name, @"^(CZ|DD|FA|HT|PP|DP)_\w+")) fileFormat = InputFileType.PatternInfo;
                    break;
                case ".txt":
                    fileFormat = GetTxtFileType(fileName); //檢查文字檔內容
                    break;
            }

            return fileFormat;
        }

        private InputFileType GetTxtFileType(string fileName)
        {
            var fileFormat = InputFileType.Unknown;

            //IGXL Sheet 通常第一行就是
            var rePatInfo = new Regex(@"^(CZ|DD|FA|HT|PP|DP)_.*\s:\s.*",
                RegexOptions.Compiled | RegexOptions.IgnoreCase); // used for Pattern Info

            //一般DataLog 只讀取第一顆Device
            var regexDebugPrint = new Regex(@"debug print start", RegexOptions.Compiled);
            var regexBcPrint = new Regex(@"_BV>", RegexOptions.Compiled);

            var regexDeviceStart = new Regex(@"Device#:", RegexOptions.Compiled);
            var regexDeviceEnd = new Regex(@"=========================================================================",
                RegexOptions.Compiled);

            var hasOneCompleteLog = 0; //頭尾都要有

            var fInfo = new FileInfo(fileName);

            var sr = new StreamReader(fInfo.FullName);
            var buffer = new char[4096];
            sr.ReadBlock(buffer, 0, buffer.Length);
            var bufferString = new string(buffer);
            if (Regex.IsMatch(bufferString,
                @"IG-XL\s+Version[\s|\S+]*Test Program[\s|\S+]*Total Execution Time[\s|\S+]*", RegexOptions.IgnoreCase))
                return InputFileType.ExecutionProfile;
            sr.Close();

            sr = new StreamReader(fInfo.FullName);
            do
            {
                var line = sr.ReadLine(); //讀取每一行
                if ((line == null) | (line == string.Empty)) continue;

                if (rePatInfo.IsMatch(line))
                {
                    fileFormat = InputFileType.PatternInfo;
                    break;
                }

                if (regexDebugPrint.IsMatch(line))
                {
                    fileFormat = InputFileType.DebugPrint;
                    break;
                }

                //特殊規格

                if (regexDeviceStart.IsMatch(line)) hasOneCompleteLog++;
                if (regexDeviceEnd.IsMatch(line))
                {
                    hasOneCompleteLog++;
                    break;
                } //要躲在Debug Mode之後
            } while (sr.Peek() != -1);

            sr.Close();

            if (hasOneCompleteLog == 2)
                //fileFormat = InputFileType.GeneralLog;

                //{
                //    fileFormat = InputFileType.DebugPrint;
                //}
                fileFormat = InputFileType.DebugPrintFromGeneralDataLog;

            if (fileFormat == InputFileType.Unknown)
            {
                sr = new StreamReader(fInfo.FullName);
                do
                {
                    var line = sr.ReadLine(); //讀取每一行
                    if ((line == null) | (line == string.Empty)) continue;
                    if (regexBcPrint.IsMatch(line))
                    {
                        fileFormat = InputFileType.GeneralLog;
                        break;
                    }
                } while (sr.Peek() != -1);

                sr.Close();
            }

            return fileFormat;
        }

        public InputTestPlan GetSelectedTestPlan()
        {
            var plans = InputFiles.FindAll(p => p is InputTestPlan && p.Selected);
            if (plans.Count >= 1)
                return plans.First() as InputTestPlan;
            return null;
        }

        public InputScgh GetSelectedScgh()
        {
            var scghFiles = InputFiles.FindAll(p => p is InputScgh && p.Selected);
            if (scghFiles.Count > 1)
                throw new Exception("More than one SCGH files has been selected!");
            if (scghFiles.Count == 0)
                //if (_inputButtonStatus.GetButtonStatus(IgnoreScghButtonName))
                return null;
            //throw new Exception("Missing SCGH file!");
            return scghFiles.First() as InputScgh;
        }

        public InputPatternListCsv GetSelectedPatternListCsv()
        {
            var patternLists = InputFiles.FindAll(p => p is InputPatternListCsv && p.Selected);
            if (patternLists.Count > 1)
                throw new Exception("More than one PatternList files has been selected!");
            if (patternLists.Count == 0)
                //if (_inputButtonStatus.GetButtonStatus(IgnorePatternListButtonName))
                //throw new Exception("Missing PatternList csv file!");
                return null;
            return patternLists.First() as InputPatternListCsv;
        }

        public InputOtpRegisterMap GetSelectedYamlFile()
        {
            var yamlFiles = InputFiles.FindAll(p =>
                p is InputOtpRegisterMap &&
                Path.GetExtension(p.FullName).Equals(".yaml", StringComparison.OrdinalIgnoreCase));
            if (yamlFiles.Count > 1)
                throw new Exception("More than one .yaml files!");
            if (yamlFiles.Count == 0)
                return null;
            return yamlFiles.First() as InputOtpRegisterMap;
        }

        public List<InputOtpRegisterMap> GetSelectedOtpFile()
        {
            var otpOptionFile = InputFiles.FindAll(p =>
                p is InputOtpRegisterMap &&
                Path.GetExtension(p.FullName).Equals(".otp", StringComparison.OrdinalIgnoreCase));
            var otpMustHaveFile = InputFiles.FindAll(p =>
                p is InputOtpRegisterMap &&
                Path.GetExtension(p.FullName).Equals(".yaml", StringComparison.OrdinalIgnoreCase));
            if (otpMustHaveFile.Count == 0)
                return null;
            if (otpOptionFile.Count != 0)
                otpMustHaveFile.AddRange(otpOptionFile);
            return otpMustHaveFile.OfType<InputOtpRegisterMap>().ToList();
        }

        public string GetCurrentProjectName(List<string> files)
        {
            var projectName = string.Empty;
            var testPlanProjectName = string.Empty;
            var scghProject = string.Empty;
            var patternListProject = string.Empty;
            var binCutProject = string.Empty;
            var vbtGenTool = string.Empty;
            foreach (var file in files)
            {
                var inputFileType = GetFileType(file);
                if (inputFileType == InputFileType.TestPlan)
                    testPlanProjectName = Path.GetFileName(file).Split('_').First();
                else if (inputFileType == InputFileType.ScghPatternList)
                    scghProject = Path.GetFileName(file).Split('_').First();
                else if (inputFileType == InputFileType.PatternListCsv)
                    patternListProject = Path.GetFileName(file).Split('_').First();
                else if (inputFileType == InputFileType.BinCut)
                    binCutProject = Path.GetFileName(file).Split('_').First();
                else if (inputFileType == InputFileType.VbtGenTool)
                    vbtGenTool = Path.GetFileName(file).Split('_').First();
            }

            if (!string.IsNullOrEmpty(testPlanProjectName))
                projectName = testPlanProjectName;

            if (!string.IsNullOrEmpty(scghProject))
            {
                if (!string.IsNullOrEmpty(projectName) &&
                    !projectName.Equals(scghProject, StringComparison.OrdinalIgnoreCase))
                    throw new Exception("The input files belong to different project!");
                if (string.IsNullOrEmpty(projectName))
                    projectName = scghProject;
            }

            if (!string.IsNullOrEmpty(patternListProject))
            {
                if (!string.IsNullOrEmpty(projectName) &&
                    !projectName.Equals(patternListProject, StringComparison.OrdinalIgnoreCase))
                    throw new Exception("The input files belong to different project!");
                if (string.IsNullOrEmpty(projectName))
                    projectName = patternListProject;
            }

            if (!string.IsNullOrEmpty(binCutProject))
            {
                if (!string.IsNullOrEmpty(projectName) &&
                    !projectName.Equals(binCutProject, StringComparison.OrdinalIgnoreCase))
                    throw new Exception("The input files belong to different project!");
                if (string.IsNullOrEmpty(projectName))
                    projectName = binCutProject;
            }

            if (string.IsNullOrEmpty(projectName))
                projectName = vbtGenTool;

            return projectName;
        }

        public string GetSelectedYamlFilePath()
        {
            var yaml = GetSelectedYamlFile();
            if (yaml == null)
                return null;
            return yaml.FullName;
        }

        public List<string> GetSelectedOtpFilePath()
        {
            var otp = GetSelectedOtpFile();
            if (otp == null)
                return new List<string>();
            return otp.Select(p => p.FullName).ToList();
        }
    }
}