using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using AutomationCommon.Controls;
using AutomationCommon.DataStructure;
using AutomationCommon.EpplusErrorReport;
using AutomationCommon.Utility;
using IgxlData.IgxlManager;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using OfficeOpenXml;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.GenerateIgxl;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.InputPackages;
using PmicAutogen.InputPackages.Base;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Inputs.PatternList;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using PmicAutogen.Properties;
using PmicAutogen.Singleton;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using Font = System.Drawing.Font;

namespace PmicAutogen
{
    public sealed partial class PmicMainForm : MyForm
    {
        //Input
        private readonly Workbook _workbook;

        private List<string> _files;

        //Output
        private InputPackage _inputPackageAutomation;

        public PmicMainForm(Workbook workbook = null)
        {
            Response.Initialize(WriteMessage);
            if (workbook != null)
            {
                _workbook = workbook;
                _workbook.Parent.DisplayAlerts = false;
                InputFiles.InteropTestPlanWorkbook = _workbook;
            }

            InitializeComponent();

            ProjectConfigSingleton.Initialize();

            myFileOpen_SettingFile.ButtonTextBox.Text = Settings.Default.SettingFile;
            if (string.IsNullOrEmpty(myFileOpen_SettingFile.ButtonTextBox.Text))
                myFileOpen_SettingFile.ButtonTextBox.Text = @"Default";
            myFileOpen_PatternPath.ButtonTextBox.Text = Settings.Default.PatthernPath;
            myFileOpen_TimeSetPath.ButtonTextBox.Text = Settings.Default.TimeSetPath;
            myFileOpen_LibraryPath.ButtonTextBox.Text = Settings.Default.LibraryPath;
            myFileOpen_ExtraPath.ButtonTextBox.Text = Settings.Default.ExtraPath;
            HelpButtonClicked += HelpButton_Clicked;

            if (_workbook != null)
                LocalSpecs.CurrentProject = Path.GetFileName(_workbook.FullName).Split('_').First();
            var assembly = Assembly.GetExecutingAssembly();
            var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            var version = fvi.FileVersion;
            Text = string.Format(@"PMIC Application ToolBox : Version - {0}", version);
        }

        private void button_LoadFiles_Click(object sender, EventArgs e)
        {
            try
            {
                CalculateTimeStart();
                Reset();
                richTextBox.Clear();
                _inputPackageAutomation = new InputPackage();
                _files = _inputPackageAutomation.OpenFileDialogWithFilter("Automation Source");
                if (_files == null)
                {
                    if (_workbook != null)
                        _files = new List<string> { _workbook.FullName };
                }
                else
                {
                    if (_workbook != null && _files.All(x => x != _workbook.FullName))
                        _files.Add(_workbook.FullName);
                }

                if (_files != null && _files.Any())
                {
                    _inputPackageAutomation.ReadFiles(_files);
                    _inputPackageAutomation.SetButtonStatus(this);
                }

                CheckStatus();
                CalculateTimeStop();
            }
            catch (Exception ex)
            {
                Response.Report(ex.Message, MessageLevel.Error, 0);
            }
        }

        private void button_RunAutogen_Click(object sender, EventArgs e)
        {
            try
            {
                if (_inputPackageAutomation.CheckInput(WriteMessage) == false)
                    goto cancelEnd;

                if (SelectOutputPath()) return;

                CalculateTimeStart();

                Initialize();

                BlockStatus.UpdateAutomationBlockStatus(this);

                SetEpWorkBook();

                var pmicGenerator = new PmicGenerator();
                pmicGenerator.Run(_workbook, _inputPackageAutomation);

                Response.Report("Generating IGXL Test Program ...", MessageLevel.CheckPoint, 0);
                var exportMain = new IgxlManagerMain();

                var igxlItems = GetAllIgxlItems();
                exportMain.GenIgxlProgram(igxlItems, LocalSpecs.TarDir, LocalSpecs.CurrentProject, TestProgram.IgxlWorkBk,
                    Response.Report, LocalSpecs.TargetIgxlVersion);

                if (EpplusErrorManager.GetErrorCount() > 0)
                    GenErrorReport();

                goto processEnd;
            cancelEnd:
                MessageBox.Show(@"User canceled or insufficient docs ...", @"Process End");
            processEnd:
                richTextBox.Enabled = true;
                CalculateTimeStop();
                SetButton(button_RunAutogen, true);
                Response.Report("Autogen is Completed", MessageLevel.EndPoint, 0);

            }
            catch (Exception ex)
            {
                Response.Report("Meet an error when running autoGen. " + ex.Message, MessageLevel.Error, 0);
            }
        }

        private static void GenErrorReport()
        {
            var app = new Application();
            if (LocalSpecs.TestPlanFileNameCopy == null ||
                LocalSpecs.TestPlanFileNameCopy.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
                return;
            var workbook = app.Workbooks.Open(LocalSpecs.TestPlanFileNameCopy, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            Response.Report("Writing Error Report ...", MessageLevel.Warning, 0);
            EpplusErrorManager.GenErrorReport(workbook, "ErrorReport");

            workbook.Save();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(app.Workbooks);
            app.Quit();
            Marshal.FinalReleaseComObject(app);
            Response.Report("Error Report is done !!!", MessageLevel.Warning, 0);
        }

        private List<string> GetAllIgxlItems()
        {
            //var mFileList = GetLibList(FolderStructure.DirLib);
            var mFileList = GetAllLibList(FolderStructure.DirModulesLib);
            mFileList.AddRange(GetLibList(Path.Combine(FolderStructure.DirLib, "PMIC")));
            mFileList.AddRange(GetLibList(FolderStructure.DirVbtGenTool));

            var mSetupFileList = mFileList.Select(nData => nData.FullName).ToList();
            mSetupFileList.AddRange(TestProgram.IgxlWorkBk.AllIgxlSheets.Keys.Select(igxlSheet => igxlSheet + ".txt"));
            mSetupFileList.AddRange(TestProgram.NonIgxlSheetsList.SheetList.Select(extraSheet => extraSheet + ".txt"));

            var igxlItems = mSetupFileList.Where(File.Exists).ToList();
            return igxlItems;
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

        private List<FileInfo> GetAllLibList(string folder)
        {
            var mFileList = new List<FileInfo>();
            if (!Directory.Exists(folder)) return mFileList;
            var dir = new DirectoryInfo(folder);
            mFileList = dir.GetFiles("*.bas", SearchOption.AllDirectories).ToList();
            mFileList.AddRange(dir.GetFiles("*.cls", SearchOption.AllDirectories).ToList());

            return mFileList;
        }

        private void SetEpWorkBook()
        {
            Response.Report(string.Format("Current Project: {0}", LocalSpecs.CurrentProject), MessageLevel.Result, 0);
            foreach (var file in _inputPackageAutomation.InputFiles)
                Response.Report(string.Format("Current {0}: {1}", file.FileType, file.FullName), MessageLevel.Result,
                    0);

            CopyInputs(_inputPackageAutomation.InputFiles);

            LocalSpecs.OtpFileName = _inputPackageAutomation.GetSelectedOtpFilePath();
            LocalSpecs.YamlFileName = _inputPackageAutomation.GetSelectedYamlFilePath();
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            if (string.IsNullOrEmpty(myFileOpen_SettingFile.ButtonTextBox.Text) ||
                myFileOpen_SettingFile.ButtonTextBox.Text.Equals("Default", StringComparison.CurrentCultureIgnoreCase))
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
                InputFiles.SettingWorkbook = new ExcelPackage(new FileInfo(myFileOpen_SettingFile.ButtonTextBox.Text)).Workbook;
            }

            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith(".Config.xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                    InputFiles.ConfigWorkbook = inputExcel.Workbook;
                    break;
                }
        }

        private void button_Setting_Click(object sender, EventArgs e)
        {
            ProjectConfigSingleton.Instance().LoadProjectConfig();
            var projectConfigSetting = new ProjectConfigSetting(LocalSpecs.CurrentProject);
            var result = projectConfigSetting.ShowDialog();
            if (result.HasValue && result.Value)
            {
                ProjectConfigSingleton.Instance().SaveProjectConfig();
                Response.Report("ProjectConfig is Saved", MessageLevel.EndPoint, 0);
            }
        }

        private void button_Clear_Click(object sender, EventArgs e)
        {
            button_Basic.Checked = false;
            button_Scan.Checked = false;
            button_Mbist.Checked = false;
            button_Otp.Checked = false;
            button_VBT.Checked = false;
        }

        #region Method

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
                    LocalSpecs.VbtGenToolFileName.Add(input.FullName);
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
            VersionControl.SrcInfoRows.Add(new SrcInfoRow(
                Text + "/ with Build Version - " + version.Major + "." + version.Minor + "." + version.Build + "." +
                version.MinorRevision, "T-AutoGen-Version"));

            foreach (var input in inputs)
                VersionControl.SrcInfoRows.Add(
                    new SrcInfoRow(input.FullName + ", MD5=" + ClassUtility.GetFileMd5(input.FullName), ""));

            #endregion

            return dic;
        }

        private void HelpButton_Clicked(object sender, EventArgs e)
        {
            var cnt = GetType().ToString().Split('.').Length;
            var name = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1)) + ".HelpFile.";
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            new MyDownloadForm().Download(name, resourceNames, assembly).Show();
        }

        public void WriteMessage(string msg, MessageLevel level = MessageLevel.General, int percentage = -1)
        {
            if (percentage >= 0)
                myStatus.ToolStripProgressBar.Value = percentage;
            if (msg.Trim().Length == 0) return;
            switch (level)
            {
                case MessageLevel.General:
                    richTextBox.SelectionColor = Color.Black;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10);
                    richTextBox.AppendText(msg + Environment.NewLine);
                    break;
                case MessageLevel.EndPoint:
                    richTextBox.SelectionColor = Color.SteelBlue;
                    richTextBox.SelectionFont = new Font("Courier New", 10, FontStyle.Bold);
                    richTextBox.AppendText("******* " + msg.PadBoth(20) + " *******" + Environment.NewLine + Environment.NewLine);
                    break;
                case MessageLevel.CheckPoint:
                    richTextBox.SelectionColor = Color.ForestGreen;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    richTextBox.AppendText(Environment.NewLine + "==================================" + Environment.NewLine);
                    richTextBox.SelectionColor = Color.ForestGreen;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    richTextBox.AppendText(msg + Environment.NewLine);
                    richTextBox.SelectionColor = Color.ForestGreen;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    richTextBox.AppendText("==================================" + Environment.NewLine + Environment.NewLine);
                    break;
                case MessageLevel.Warning:
                    richTextBox.SelectionColor = Color.Coral;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10);
                    richTextBox.AppendText("[Warning] " + msg + Environment.NewLine);
                    break;
                case MessageLevel.Result:
                    richTextBox.SelectionColor = Color.ForestGreen;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    richTextBox.AppendText(msg + Environment.NewLine);
                    break;
                case MessageLevel.Error:
                    richTextBox.SelectionColor = Color.Red;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
                    richTextBox.AppendText("[Error] " + msg + Environment.NewLine);
                    break;
                default:
                    richTextBox.SelectionColor = Color.Black;
                    richTextBox.SelectionFont = new Font("Microsoft Sans Serif", 10);
                    richTextBox.AppendText(msg + Environment.NewLine);
                    break;
            }

            Thread.Sleep(10);
            richTextBox.ScrollToCaret(); // scroll to the end
            richTextBox.Refresh();
        }

        private void Reset()
        {
            richTextBox.Clear();
            myStatus.ToolStripStatusLabel.Text = @"Status";
            myStatus.ProcessTimeToolStripStatusLabel.Text = @"Process Time";
        }

        private void CalculateTimeStart()
        {
            StartTime = DateTime.Now;
            myStatus.ToolStripProgressBar.Value = 0;
            myStatus.ProcessTimeToolStripStatusLabel.Text = @"Process Time";
        }

        private void CalculateTimeStop()
        {
            myStatus.ToolStripStatusLabel.Text = @"Done";
            myStatus.ProcessTimeToolStripStatusLabel.Text = (DateTime.Now - StartTime).ToString(@"hh\:mm\:ss");
        }

        private bool SelectOutputPath()
        {
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = Settings.Default.OutputPath,
                Title = @"Select Directory of Setting Folder"
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
                return true;

            LocalSpecs.TarDir = folderBrowserDialog.FileName;
            LocalSpecs.PatternPath = myFileOpen_PatternPath.ButtonTextBox.Text;
            LocalSpecs.TimeSetPath = myFileOpen_TimeSetPath.ButtonTextBox.Text;
            LocalSpecs.BasLibraryPath = myFileOpen_LibraryPath.ButtonTextBox.Text;
            LocalSpecs.SettingFile = myFileOpen_SettingFile.ButtonTextBox.Text;
            LocalSpecs.ExtraPath = myFileOpen_ExtraPath.ButtonTextBox.Text;

            FolderStructure.ResetFolderVaribles();

            Settings.Default.OutputPath = folderBrowserDialog.FileName;
            Settings.Default.Save();
            return false;
        }

        private void Initialize()
        {
            SetButton(button_RunAutogen, false);
            TestProgram.Initialize();
            LocalSpecs.Initialize();
            InputFiles.Initialize();
            EpplusErrorManager.Initialize();
            BinNumberSingleton.Initialize();
            SheetStructureManager.Initialize();
            VersionControl.Initialize();
            HardIpDataMain.Initialize();
            CharSetupSingleton.Initialize();
            if (Directory.Exists(FolderStructure.DirTrunk))
                Directory.Delete(FolderStructure.DirTrunk, true);

            richTextBox.Clear();
        }

        private void FileOpen_SettingFile_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            FileSelect(sender, FileFilter.Excel);
        }

        private void myFileOpen_PatternPath_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            PathSelect(sender);
        }

        private void myFileOpen_TimeSetPath_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            PathSelect(sender);
        }

        private void myFileOpen_LibraryPath_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            PathSelect(sender);
        }

        private void myFileOpen_ExtraPath_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            PathSelect(sender);
        }

        private void CheckStatus()
        {
            var flag = true;
            if (string.IsNullOrEmpty(myFileOpen_SettingFile.ButtonTextBox.Text))
                flag = false;
            else
                Settings.Default.SettingFile = myFileOpen_SettingFile.ButtonTextBox.Text;

            if (string.IsNullOrEmpty(myFileOpen_PatternPath.ButtonTextBox.Text))
                flag = false;
            else
                Settings.Default.PatthernPath = myFileOpen_PatternPath.ButtonTextBox.Text;

            if (string.IsNullOrEmpty(myFileOpen_TimeSetPath.ButtonTextBox.Text))
                flag = false;
            else
                Settings.Default.TimeSetPath = myFileOpen_TimeSetPath.ButtonTextBox.Text;

            if (string.IsNullOrEmpty(myFileOpen_LibraryPath.ButtonTextBox.Text))
                flag = false;
            else
                Settings.Default.LibraryPath = myFileOpen_LibraryPath.ButtonTextBox.Text;


            if (string.IsNullOrEmpty(myFileOpen_ExtraPath.ButtonTextBox.Text))
                flag = false;
            else
                Settings.Default.ExtraPath = myFileOpen_ExtraPath.ButtonTextBox.Text;

            if (_files == null || !_files.Any())
                flag = false;

            Settings.Default.Save();
            SetButton(button_RunAutogen, flag);
        }

        private void SetButton(Button button, bool flag)
        {
            button.Enabled = flag;
            button.ForeColor = flag ? SystemColors.ActiveCaptionText : SystemColors.InactiveCaptionText;
        }

        private void FileOpen_SettingFile_ButtonTextBoxTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        private void myFileOpen_PatternPath_ButtonTextBoxTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        private void myFileOpen_TimeSetPath_ButtonTextBoxTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        private void myFileOpen_LibraryPath_ButtonTextBoxTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        private void myFileOpen_ExtraPath_ButtonTextBoxTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        #endregion
    }
}