using CommonLib.Enum;
using CommonLib.WriteMessage;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic;
using PmicAutogen.GenerateIgxl.Mbist;
using PmicAutogen.GenerateIgxl.OTP;
using PmicAutogen.GenerateIgxl.PostAction;
using PmicAutogen.GenerateIgxl.PreAction;
using PmicAutogen.GenerateIgxl.Scan;
using PmicAutogen.GenerateIgxl.VbtGenTool;
using PmicAutogen.GenerateIgxl.VbtGenTool.Checker;
using PmicAutogen.InputPackages.Base;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Inputs.CopyXml;
using PmicAutogen.Inputs.OtpFiles;
using PmicAutogen.Inputs.PatternList;
using PmicAutogen.Inputs.ScghFile;
using PmicAutogen.Inputs.Setting;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.VbtGenTool;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using PmicAutogen.UI;
using PmicAutogen.ViewModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PmicAutogen.GenerateIgxl
{
    public class PmicGenerator
    {
        public PmicGenerator(List<Input> inputFiles)
        {
            SetEpWorkBook();

            CopyInputs(inputFiles);
        }

        private void SetEpWorkBook()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            if (string.IsNullOrEmpty(LocalSpecs.SettingFile) ||
                LocalSpecs.SettingFile.Equals("Default", StringComparison.CurrentCultureIgnoreCase))
            {
                foreach (var resourceName in resourceNames)
                    if (resourceName.EndsWith(".Setting.xlsx", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                        InputFiles.SettingWorkbook = inputExcel.Workbook;
                        break;
                    }
            }
            else
            {
                InputFiles.SettingWorkbook = new ExcelPackage(new FileInfo(LocalSpecs.SettingFile)).Workbook;
            }

            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith(".Config.xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                    InputFiles.ConfigWorkbook = inputExcel.Workbook;
                    break;
                }
        }

        public Dictionary<Input, string> CopyInputs(List<Input> inputs)
        {
            var dic = new Dictionary<Input, string>();
            var exist = Directory.Exists(FolderStructure.DirIgLink);
            if (!exist)
                Directory.CreateDirectory(FolderStructure.DirIgLink);

            foreach (var input in inputs)
            {
                var file = input.FullName;
                var extension = Path.GetExtension(file);

                if (input.FileType == InputFileType.TestPlan)
                {
                    LocalSpecs.TestPlanFileName = input.FullName;
                }
                else if (input.FileType == InputFileType.ScghPatternList)
                {
                    LocalSpecs.ScghFileName = input.FullName;
                }
                else if (input.FileType == InputFileType.VbtGenTool)
                {
                    LocalSpecs.VbtGenToolFileNames.Add(input.FullName);
                }
                else if (input.FileType == InputFileType.PatternListCsv)
                {
                    LocalSpecs.PatListCsvFile = input.FullName;
                    InputFiles.PatternListMap = PatternListMap.Initialize(LocalSpecs.PatListCsvFile,
                        LocalSpecs.TimeSetPath, LocalSpecs.PatternPath);
                }

                if (extension != null && !extension.StartsWith(".xls", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                var sourcePath = file;
                var targetFile = VersionControl.AddTimeStamp(file);
                var targetPath = Path.Combine(FolderStructure.DirIgLink, targetFile);
                if (!Directory.Exists(FolderStructure.DirIgLink))
                    Directory.CreateDirectory(FolderStructure.DirIgLink);
                if (File.Exists(targetPath))
                    File.Delete(targetPath);
                if (sourcePath != null) File.Copy(sourcePath, targetPath);
                dic.Add(input, targetPath);
            }

            foreach (var item in dic)
            {
                var inputExcel = new ExcelPackage(new FileInfo(item.Value));
                if (item.Key.FileType == InputFileType.TestPlan)
                {
                    InputFiles.TestPlanExcelPackage = inputExcel;
                    InputFiles.TestPlanWorkbook = inputExcel.Workbook;
                    LocalSpecs.TestPlanFileNameCopy = item.Value;
                }
                else if (item.Key.FileType == InputFileType.ScghPatternList)
                {
                    InputFiles.ScghPackage = inputExcel;
                    InputFiles.ScghWorkbook = inputExcel.Workbook;
                    LocalSpecs.ScghFileNameCopy = item.Value;
                }
                else if (item.Key.FileType == InputFileType.VbtGenTool)
                {
                    InputFiles.VbtGenToolPackage.Add(inputExcel);
                    InputFiles.VbtGenToolWorkbooks.Add(inputExcel.Workbook);
                    LocalSpecs.VbtGenToolFileNameCopy.Add(item.Value);
                }
            }

            #region set input version

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            VersionControl.SrcInfoRows.Add(new SrcInfoRow("T-AutoGen-Version", version.ToString()));

            foreach (var input in inputs)
                VersionControl.SrcInfoRows.Add(new SrcInfoRow(Path.GetFileName(input.FullName), ""));

            #endregion

            return dic;
        }

        public void Run(Workbook workbook)
        {
            ReadFile();

            PostCheck();

            GenProgram();
        }

        private void GenProgram()
        {
            try
            {
                #region PreAction
                Response.Report("Running Pre-Action ...", EnumMessageLevel.CheckPoint, 0);
                using (var preActionMain = new PreActionMain())
                {
                    preActionMain.WorkFlow();
                }
                Response.Report("Pre-Action is Completed", EnumMessageLevel.EndPoint, 0);
                #endregion

                #region Basic
                if (ViewModelMain.Instance().BasicIsChecked)
                {
                    Response.Report("Running Basic ...", EnumMessageLevel.CheckPoint, 0);
                    using (var basicMain = new BasicMain())
                    {
                        basicMain.WorkFlow();
                    }

                    Response.Report("Basic is Completed", EnumMessageLevel.EndPoint, 0);
                }
                #endregion

                #region Scan
                if (ViewModelMain.Instance().ScanIsChecked)
                {
                    Response.Report("Running Scan ...", EnumMessageLevel.CheckPoint, 0);
                    using (var scanMain = new ScanMain())
                    {
                        scanMain.WorkFlow();
                    }

                    Response.Report("Scan is Completed", EnumMessageLevel.EndPoint, 0);
                }
                #endregion

                #region Mbist
                if (ViewModelMain.Instance().MbistIsChecked)
                {
                    Response.Report("Running Mbist ...", EnumMessageLevel.CheckPoint, 0);
                    using (var mbistMain = new MbistMain())
                    {
                        mbistMain.WorkFlow();
                    }

                    Response.Report("Mbist is Completed", EnumMessageLevel.EndPoint, 0);
                }
                #endregion

                #region OTP
                if (ViewModelMain.Instance().OTPIsChecked)
                {
                    Response.Report("Running OTP ...", EnumMessageLevel.CheckPoint, 0);
                    using (var otpMain = new OtpMain())
                    {
                        otpMain.WorkFlow();
                    }

                    Response.Report("OTP is Completed", EnumMessageLevel.EndPoint, 0);
                }
                #endregion

                #region VBT
                if (ViewModelMain.Instance().VBTIsChecked)
                    foreach (var testParameterSheet in StaticVbtGenTool.TestParameterSheets)
                    {
                        Response.Report(string.Format("Running {0} ...", testParameterSheet.Block), EnumMessageLevel.CheckPoint,
                            0);
                        using (var main = new VbtGenMain(testParameterSheet))
                        {
                            main.WorkFlow();
                        }

                        Response.Report(testParameterSheet.Block + " Generation Completed", EnumMessageLevel.EndPoint, 0);
                    }
                #endregion

                #region PostAction
                Response.Report("Running Post-Action ...", EnumMessageLevel.CheckPoint, 0);
                using (var postActionMain = new PostActionMain())
                {
                    postActionMain.WorkFlow();
                    Response.Report("Exporting ASCII txt files ...", EnumMessageLevel.General, 100);
                    postActionMain.PrintFlow();
                }

                Response.Report("Post-Action is Completed", EnumMessageLevel.EndPoint, 0);
                #endregion
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in generating progrma. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        private void ReadFile()
        {
            try
            {
                if (InputFiles.SettingWorkbook != null)
                {
                    Response.Report("Reading Setting ...", EnumMessageLevel.General, 0);
                    var settingManager = new SettingManager();
                    settingManager.CheckAll(InputFiles.SettingWorkbook);
                    StaticSetting.AddSheets(settingManager);
                }

                if (InputFiles.TestPlanWorkbook != null)
                {
                    Response.Report("Reading TestPlan ...", EnumMessageLevel.General, 0);
                    var testPlanManager = new TestPlanManager();
                    testPlanManager.ReadAll(InputFiles.TestPlanWorkbook);
                    StaticTestPlan.AddSheets(testPlanManager);
                }

                #region vddRefForm
                var domainDic = StaticTestPlan.IoLevelsSheet.Rows.GroupBy(x => x.Domain)
                    .ToDictionary(x => x.Key, y => y.First().IoLevelDate.First().Vdd);
                var vddPinDic = StaticTestPlan.VddLevelsSheet.Rows
                     .ToDictionary(x => x.WsBumpName, y => y);

                if (!LocalSpecs.IsUnitTest)
                {
                    var vddRefWindow = new VDDRefWindow(domainDic, vddPinDic);
                    if ((bool)vddRefWindow.ShowDialog())
                        foreach (var pin in vddRefWindow.RefVddPins)
                            LocalSpecs.VddRefInfoList.Add(pin.Key, pin.Value);
                }
                else
                {
                    LocalSpecs.VddRefInfoList = new Dictionary<string, VddLevelsRow>();
                }


                //var vddRefForm = new VddRefForm(domainDic, vddPinDic);
                //if (!LocalSpecs.IsUnitTest)
                //{
                //    if (vddRefForm.ShowDialog() == DialogResult.OK)
                //        foreach (var pin in vddRefForm.RefVddPins)
                //            LocalSpecs.VddRefInfoList.Add(pin.Key, pin.Value);
                //}
                //else
                //{
                //    vddRefForm.Click_Ok();
                //    foreach (var pin in vddRefForm.RefVddPins)
                //        LocalSpecs.VddRefInfoList.Add(pin.Key, pin.Value);
                //}
                #endregion

                if (InputFiles.ScghWorkbook != null)
                {
                    Response.Report("Reading SCGH file ...", EnumMessageLevel.General, 0);
                    var scghFileManager = new ScghFileManager();
                    scghFileManager.CheckAll(InputFiles.ScghWorkbook);
                    StaticScgh.AddSheets(scghFileManager);
                }

                if (!string.IsNullOrEmpty(LocalSpecs.YamlFileName))
                {
                    Response.Report("Reading OTP Files ...", EnumMessageLevel.General, 0);
                    var otpManager = new OtpManager();
                    otpManager.CheckAll();
                    StaticOtp.AddSheets(otpManager);
                }

                if (InputFiles.VbtGenToolWorkbooks != null && InputFiles.VbtGenToolWorkbooks.Any())
                {
                    Response.Report("Reading VbtGenTool Files ...", EnumMessageLevel.General, 0);
                    var vbtGenToolManager = new VbtGenToolManager();
                    vbtGenToolManager.CheckAll(InputFiles.VbtGenToolWorkbooks);
                    StaticVbtGenTool.AddSheets(vbtGenToolManager);
                }

                var copyNwire = new CopyXmlFiles();
                copyNwire.Work();
                var copyConfig = new CopyIgxlConfigFiles();
                copyConfig.Work();
            }
            catch (Exception e)
            {
                Response.Report("Meet an error when reading files. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        private void PostCheck()
        {
            try
            {
                new PinNameChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets);

                new BitFieldChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets, StaticTestPlan.AhbRegisterMapSheet);

                new AhbRegisterChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets,
                    StaticTestPlan.AhbRegisterMapSheet);

                new PatternChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in post check. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        public List<string> GetAllIgxlItems()
        {
            //var mFileList = GetLibList(FolderStructure.DirLib);
            var mFileList = GetAllLibList(FolderStructure.DirModulesLibTer);
            mFileList.AddRange(GetLibList(Path.Combine(FolderStructure.DirOtherWaitForClassify, "PMIC")));
            mFileList.AddRange(GetLibList(FolderStructure.DirOtp));

            var mSetupFileList = mFileList.Select(nData => nData.FullName).ToList();
            mSetupFileList.AddRange(TestProgram.IgxlWorkBk.AllIgxlSheets.Keys.Select(igxlSheet => igxlSheet + ".txt"));
            mSetupFileList.AddRange(TestProgram.NonIgxlSheetsList.SheetList.Select(extraSheet => extraSheet + ".txt"));

            var igxlItems = mSetupFileList.Where(File.Exists).ToList();
            return igxlItems;
        }

        private List<FileInfo> GetAllLibList(string folder)
        {
            var mFileList = new List<FileInfo>();
            if (!Directory.Exists(folder)) return mFileList;
            var dir = new DirectoryInfo(folder);
            mFileList = dir.GetFiles("*.bas", SearchOption.AllDirectories).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.AllDirectories).ToList());

            return mFileList;
        }

        private List<FileInfo> GetLibList(string folder)
        {
            var mFileList = new List<FileInfo>();
            if (!Directory.Exists(folder)) return mFileList;
            var dir = new DirectoryInfo(folder);
            mFileList = dir.GetFiles("*.bas", SearchOption.TopDirectoryOnly).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.TopDirectoryOnly).ToList());
            foreach (var subDir in dir.GetDirectories())
            {
                mFileList.AddRange(subDir.GetFiles("*.bas", SearchOption.TopDirectoryOnly));
                mFileList.AddRange(subDir.GetFiles("*.cls", SearchOption.TopDirectoryOnly));
            }

            return mFileList;
        }
    }
}