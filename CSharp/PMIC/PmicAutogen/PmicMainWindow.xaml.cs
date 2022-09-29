using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using IgxlData.IgxlManager;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using MyWpf.Controls;
using OfficeOpenXml;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.GenerateIgxl;
using PmicAutogen.InputPackages;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using PmicAutogen.Local.Version;
using PmicAutogen.Properties;
using PmicAutogen.Singleton;
using PmicAutogen.ViewModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PmicAutogen
{
    public partial class PmicMainWindow : MyWindow
    {
        //Input
        private Workbook _workbook;
        //Output
        private InputPackage _inputPackageAutomation = new InputPackage();

        public PmicMainWindow()
        {
            PmicMainWindowInitialize(null);
        }

        public PmicMainWindow(Workbook workbook = null)
        {
            PmicMainWindowInitialize(workbook);
        }

        private void PmicMainWindowInitialize(Workbook workbook)
        {
            HelpButtonClick += HelpButton_Clicked;

            InitializeComponent();

            DataContext = ViewModelMain.Instance();

            LoadSettings();

            Response.Progress = WriteMessage();

            if (workbook != null)
            {
                _workbook = workbook;
                CheckBtnSettingIsEnabled(workbook);
                InputFiles.InteropTestPlanWorkbook = _workbook;
                _inputPackageAutomation.ReadFiles(new List<string>() { _workbook.FullName });
                SetButtonStatus(_inputPackageAutomation);
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = true;
            }
            else
            {
                ViewModelMain.Instance().BtnSettingIsEnabled = false;
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            }

            InitializeComponent();

            ProjectConfigSingleton.Initialize();

            if (_workbook != null)
                LocalSpecs.CurrentProject = Path.GetFileName(_workbook.FullName).Split('_').First();

            var assembly = Assembly.GetExecutingAssembly();
            var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            var version = fvi.FileVersion;
            Title += string.Format(@"Version - {0}", version);
        }

        private async void button_LoadFiles_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.Wait;
                ((Control)sender).IsEnabled = false;
                TimeStart();
                ClearUI();

                _inputPackageAutomation.Clear(_workbook);
                var files = _inputPackageAutomation.OpenFileDialogWithFilter("Automation Source");
                await Task.Factory.StartNew(() =>
                {
                    if (files != null && files.Any())
                    {
                        if (_workbook != null)
                            files.Remove(_workbook.FullName);
                        _inputPackageAutomation.ReadFiles(files);
                        SetButtonStatus(_inputPackageAutomation);
                    }
                });
                CheckStatus();
                TimeStop();
            }
            catch (Exception ex)
            {
                Response.Report(ex.Message, EnumMessageLevel.Error, 0);
            }
            finally
            {
                ((Control)sender).IsEnabled = true;
                Cursor = null;
            }
        }

        private async void button_RunAutogen_Click(object sender, EventArgs e)
        {

            try
            {
                Cursor = Cursors.Wait;
                ClearUI();

                if (_inputPackageAutomation.CheckInput(WriteMessage) == false)
                    return;

                if (SelectOutputPath())
                {
                    Response.Report(@"User canceled porcess ...", EnumMessageLevel.Error, 0);
                    return;
                }

                TimeStart();

                Initialize();

                //BlockStatus.UpdateAutomationBlockStatus(
                //    MainViewModel.Instance().BasicIsChecked,
                //    MainViewModel.Instance().ScanIsChecked,
                //    MainViewModel.Instance().MbistIsChecked,
                //    MainViewModel.Instance().OTPIsChecked,
                //    MainViewModel.Instance().VBTIsChecked,
                //    MainViewModel.Instance().BasicIsEnabled,
                //    MainViewModel.Instance().ScanIsEnabled,
                //    MainViewModel.Instance().MbistIsEnabled,
                //    MainViewModel.Instance().OTPIsEnabled,
                //    MainViewModel.Instance().VBTIsEnabled);

                Response.Report(string.Format("Current Project: {0}", LocalSpecs.CurrentProject), EnumMessageLevel.Result, 0);
                foreach (var file in _inputPackageAutomation.InputFiles)
                    Response.Report(string.Format("Current {0}: {1}", file.FileType, file.FullName), EnumMessageLevel.Result, 0);

                LocalSpecs.OtpFileNames = _inputPackageAutomation.GetSelectedOtpFilePath();
                LocalSpecs.YamlFileName = _inputPackageAutomation.GetSelectedYamlFilePath();

                var taskScheduler = new StaTaskScheduler(1);
                await Task.Factory.StartNew(() =>
                {
                    var pmicGenerator = new PmicGenerator(_inputPackageAutomation.InputFiles);
                    pmicGenerator.Run(_workbook);
                    var igxlItems = pmicGenerator.GetAllIgxlItems();

                    Response.Report("Generating IGXL Test Program ...", EnumMessageLevel.CheckPoint, 0);
                    var exportMain = new IgxlManagerMain();
                    exportMain.GenIgxlProgram(igxlItems, LocalSpecs.TarDir, LocalSpecs.CurrentProject,
                        TestProgram.IgxlWorkBk, Response.Report, LocalSpecs.TargetIgxlVersion);
                }, CancellationToken.None, TaskCreationOptions.None, taskScheduler);
                //.ContinueWith((t) =>
                //{
                //    MessageBox.Show(t.Exception.ToString());
                //}, TaskContinuationOptions.OnlyOnFaulted);

                if (ErrorManager.GetErrorCount() > 0)
                    GenErrorReport();

                TimeStop();
                Response.Report("Autogen is Completed", EnumMessageLevel.EndPoint, 0);
            }
            catch (Exception ex)
            {
                Response.Report("Meet an error when running autoGen. " + ex.Message, EnumMessageLevel.Error, 0);
            }
            finally
            {
                Cursor = null;
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = true;
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
                Response.Report("ProjectConfig is Saved", EnumMessageLevel.EndPoint, 0);
            }
        }

        private void button_Clear_Click(object sender, EventArgs e)
        {
            button_Basic.IsChecked = false;
            button_Scan.IsChecked = false;
            button_Mbist.IsChecked = false;
            button_OTP.IsChecked = false;
            button_VBT.IsChecked = false;
        }

        #region Method
        public void GenErrorReport()
        {
            var app = new Application();
            if (LocalSpecs.TestPlanFileNameCopy == null ||
                LocalSpecs.TestPlanFileNameCopy.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
                return;
            var workbook = app.Workbooks.Open(LocalSpecs.TestPlanFileNameCopy, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            try
            {
                ErrorManager.GenErrorReport(workbook, "ErrorReport");
                Response.Report("Writing Error Report ...", EnumMessageLevel.Error, 0);
                Response.Report(workbook.FullName, EnumMessageLevel.Error, 0);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(workbook);
                Marshal.FinalReleaseComObject(app.Workbooks);
                app.Quit();
                Marshal.FinalReleaseComObject(app);
            }
        }

        public void SetButtonStatus(InputPackage _inputPackageAutomation)
        {
            ViewModelMain.Instance().SetButtonStatusTrue();

            ViewModelMain.Instance().BasicIsEnabled = _inputPackageAutomation.HasTestPlan;
            ViewModelMain.Instance().ScanIsEnabled = _inputPackageAutomation.HasScan;
            ViewModelMain.Instance().MbistIsEnabled = _inputPackageAutomation.HasMbist;
            ViewModelMain.Instance().OTPIsEnabled = _inputPackageAutomation.HasOtpFile;
            ViewModelMain.Instance().VBTIsEnabled = _inputPackageAutomation.HasVbtFile;

            ViewModelMain.Instance().BasicIsChecked = _inputPackageAutomation.HasTestPlan;
            ViewModelMain.Instance().ScanIsChecked = _inputPackageAutomation.HasScan;
            ViewModelMain.Instance().MbistIsChecked = _inputPackageAutomation.HasMbist;
            ViewModelMain.Instance().OTPIsChecked = _inputPackageAutomation.HasOtpFile;
            ViewModelMain.Instance().VBTIsChecked = _inputPackageAutomation.HasVbtFile;

            if (InputFiles.InteropTestPlanWorkbook != null)
            {
                Workbook workbook = InputFiles.InteropTestPlanWorkbook;
                CheckBtnSettingIsEnabled(workbook);
            }
            else
            {
                if (_inputPackageAutomation.InputFiles.Any(x => x.FileType == InputFileType.TestPlan))
                {
                    var file = _inputPackageAutomation.InputFiles.First(x => x.FileType == InputFileType.TestPlan).FullName;
                    var inputExcel = new ExcelPackage(new FileInfo(file));
                    InputFiles.TestPlanWorkbook = inputExcel.Workbook;
                    var sheet = inputExcel.Workbook.Worksheets[PmicConst.ProjectConfig];
                    if (sheet != null)
                        ViewModelMain.Instance().BtnSettingIsEnabled = true;
                }
            }
        }

        private static void CheckBtnSettingIsEnabled(Workbook workbook)
        {
            Worksheet worksheet = workbook.GetSheet(PmicConst.ProjectConfig);
            if (worksheet != null)
                ViewModelMain.Instance().BtnSettingIsEnabled = true;
        }

        private void HelpButton_Clicked(object sender, EventArgs e)
        {
            var cnt = GetType().ToString().Split('.').Length;
            var name = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1)) + ".HelpFiles.";
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            var helpWindow = new HelpWindow().Download(name, resourceNames, assembly);
            helpWindow.ShowDialog();
        }

        private Progress<ProgressStatus> WriteMessage()
        {
            var progress = new Progress<ProgressStatus>();
            progress.ProgressChanged += (o, info) =>
            {
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action<string, EnumMessageLevel, int>(WriteMessage)
                     , info.Message, info.Level, info.Percentage);
                }
                else
                {
                    WriteMessage(info.Message, info.Level, info.Percentage);
                }
            };
            return progress;
        }

        private void WriteMessage(string msg, EnumMessageLevel level = EnumMessageLevel.General, int percentage = -1)
        {
            if (percentage >= 0)
                MyStatusBar.ProgressBarValue = percentage;
            if (msg.Trim().Length == 0) return;

            var fontSize = 14;

            switch (level)
            {
                case EnumMessageLevel.General:
                    {
                        Run r = new Run(msg);
                        r.Foreground = new SolidColorBrush(Colors.Black);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
                case EnumMessageLevel.EndPoint:
                    {
                        Run r = new Run("******* " + msg.PadBoth(20) + " *******" + Environment.NewLine);
                        r.Foreground = new SolidColorBrush(Colors.SteelBlue);
                        r.FontFamily = new FontFamily("Courier New");
                        r.FontSize = fontSize + 6;
                        r.FontWeight = FontWeights.Bold;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
                case EnumMessageLevel.CheckPoint:
                    {
                        Run r = new Run(Environment.NewLine + "==================================");
                        r.Foreground = new SolidColorBrush(Colors.ForestGreen);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize;
                        r.FontWeight = FontWeights.Bold;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        Run r1 = new Run(msg);
                        r1.Foreground = new SolidColorBrush(Colors.ForestGreen);
                        r1.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r1.FontSize = fontSize;
                        r1.FontWeight = FontWeights.Bold;
                        Paragraph paragraph1 = new Paragraph(r1);
                        MyRichTextBox.Document.Blocks.Add(paragraph1);
                        Run r2 = new Run("==================================");
                        r2.Foreground = new SolidColorBrush(Colors.ForestGreen);
                        r2.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r2.FontSize = fontSize;
                        r2.FontWeight = FontWeights.Bold;
                        Paragraph paragraph2 = new Paragraph(r2);
                        MyRichTextBox.Document.Blocks.Add(paragraph2);
                        break;
                    }
                case EnumMessageLevel.Warning:
                    {
                        Run r = new Run("[Warning] " + msg);
                        r.Foreground = new SolidColorBrush(Colors.Coral);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
                case EnumMessageLevel.Result:
                    {
                        Run r = new Run(msg);
                        r.Foreground = new SolidColorBrush(Colors.ForestGreen);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
                case EnumMessageLevel.Error:
                    {
                        Run r = new Run("[Error] " + msg);
                        r.Foreground = new SolidColorBrush(Colors.Red);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize + 2;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
                default:
                    {
                        Run r = new Run(msg);
                        r.Foreground = new SolidColorBrush(Colors.Black);
                        r.FontFamily = new FontFamily("Microsoft Sans Serif");
                        r.FontSize = fontSize;
                        Paragraph paragraph = new Paragraph(r);
                        MyRichTextBox.Document.Blocks.Add(paragraph);
                        break;
                    }
            }

            Thread.Sleep(10);
            MyRichTextBox.ScrollToEnd();
        }

        private void ClearUI()
        {
            MyRichTextBox.Document.Blocks.Clear();
            MyStatusBar.StatusText = "";
            MyStatusBar.ProcessTimeText = "";
        }

        private void TimeStart()
        {
            StartTime = TimeProvider.Current.Now;
            MyStatusBar.ProgressBarValue = 0;
            MyStatusBar.ProcessTimeText = "";
        }

        private void TimeStop()
        {
            MyStatusBar.StatusText = "Done";
            MyStatusBar.ProcessTimeText = string.Format("Process time : {0}", (TimeProvider.Current.Now - StartTime).ToString(@"hh\:mm\:ss"));
        }

        private bool SelectOutputPath()
        {
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = Settings.Default.OutputPath,
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
                return true;

            LocalSpecs.TarDir = folderBrowserDialog.FileName;
            LocalSpecs.PatternPath = PatternPath.Text;
            LocalSpecs.TimeSetPath = TimeSetPath.Text;
            LocalSpecs.BasLibraryPath = LibraryPath.Text;
            LocalSpecs.SettingFile = Setting.Text;
            LocalSpecs.ExtraPath = ExtraSheetPath.Text;

            //FolderStructure.ResetFolderVaribles();

            Settings.Default.OutputPath = folderBrowserDialog.FileName;
            Settings.Default.Save();
            return false;
        }

        private void Initialize()
        {
            ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            LocalSpecs.Initialize();
            TestProgram.Initialize();
            InputFiles.Initialize();
            ErrorManager.Initialize();
            BinNumberSingleton.Initialize();
            SheetStructureManager.Initialize();
            VersionControl.Initialize();
            CharSetupSingleton.Initialize();
        }

        private void Setting_Click(object sender, EventArgs e)
        {
            FileSelect(sender, EnumFileFilter.Excel);
            CheckStatus();
        }

        private void PatternPath_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void TimeSetPath_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void LibraryPath_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void ExtraSheetPath_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void CheckStatus()
        {
            if (string.IsNullOrEmpty(Setting.Text))
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            else
                Settings.Default.SettingFile = Setting.Text;

            if (string.IsNullOrEmpty(PatternPath.Text))
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            else
                Settings.Default.PatternPath = PatternPath.Text;

            if (string.IsNullOrEmpty(TimeSetPath.Text))
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            else
                Settings.Default.TimeSetPath = TimeSetPath.Text;

            if (string.IsNullOrEmpty(LibraryPath.Text))
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            else
                Settings.Default.LibraryPath = LibraryPath.Text;

            if (string.IsNullOrEmpty(ExtraSheetPath.Text))
                ViewModelMain.Instance().BtnRunAutogenIsEnabled = false;
            else
                Settings.Default.ExtraPath = ExtraSheetPath.Text;

            Settings.Default.Save();
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

        private void LoadSettings()
        {
            Setting.Text = Settings.Default.SettingFile;
            if (string.IsNullOrEmpty(Setting.Text))
                Setting.Text = @"Default";
            PatternPath.Text = Settings.Default.PatternPath;
            TimeSetPath.Text = Settings.Default.TimeSetPath;
            LibraryPath.Text = Settings.Default.LibraryPath;
            ExtraSheetPath.Text = Settings.Default.ExtraPath;

            CheckStatus();
        }
        #endregion
    }
}
