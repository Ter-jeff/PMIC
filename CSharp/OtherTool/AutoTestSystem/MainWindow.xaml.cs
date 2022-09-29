using AutoTestSystem.Function;
using AutoTestSystem.Model;
using AutoTestSystem.Properties;
using AutoTestSystem.UI;
using AutoTestSystem.ViewModel;
using CommonLib.Enum;
using CommonLib.Utility;
using Microsoft.WindowsAPICodePack.Dialogs;
using MyWpf.Controls;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.Forms.MessageBox;

namespace AutoTestSystem
{
    public partial class MainWindow
    {
        private const string SettingIni = "Setting.ini";

        //private const string DefaultInput = @"U:\TP-to-C651\PatternValidation\Input";
        private const string TesterStatusReport = @"C:"; //@"U:\TP-to-C651\PatternValidation";
        private const string TesterStatusFileName = @"TesterStatus.xlsx";

        //private const string RecordTableName = "TesterStatus.db";

        private static NotifyIcon _trayIcon;

        private readonly ConcurrentQueue<QueueFile> _queueFiles = new ConcurrentQueue<QueueFile>();
        //private readonly string _testerStatusdb = Path.Combine(TesterStatusReport, RecordTableName);
        private readonly string _testerStatusExcel = Path.Combine(TesterStatusReport, TesterStatusFileName);
        private readonly BackgroundWorker _worker = new BackgroundWorker();
        private int _maxWaitTime;
        private int _period;

        private FileSystemWatcher _watcher;

        public ViewModelMain MyViewModel = new ViewModelMain();

        public MainWindow()
        {
            HelpButtonClick += HelpButton_Clicked;

            InitializeComponent();

            _worker.DoWork += Worker_DoWork;

            if (Settings.Default == null)
                return;

            WatchFolder.Text = Settings.Default.WatchFolder ?? "";

            CheckStatus();

            WatchingButton.IsEnabled = !HasExcel();

            EnvCheck();
        }

        private void EnvCheck()
        {
            if (!Directory.Exists(TesterStatusReport))
                MessageBox.Show(string.Format(@"Please check if this tester can connect net disk {0} !!!",
                    TesterStatusReport));
            if (IsOpened(_testerStatusExcel))
                MessageBox.Show(string.Format(@"Please check if the file {0} is opened !!!", _testerStatusExcel));
        }

        public bool IsOpened(string filePath)
        {
            if (!File.Exists(filePath)) return false;
            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        private void Button_WatchFolder_Click(object sender, EventArgs e)
        {
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                Title = @"Select Monitor Folder",
                IsFolderPicker = true,
                Multiselect = false
            };
            if (WatchFolder.Text != "")
                folderBrowserDialog.InitialDirectory = WatchFolder.Text;

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel) return;
            WatchFolder.Text = folderBrowserDialog.FileName;
            MyDataGrid.ItemsSource = null;
            CheckStatus();
        }

        private void ProjectComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MyViewModel.SelectProject == null) return;
            var dirs = Directory.GetDirectories(MyViewModel.SelectProject.Path, "*", SearchOption.TopDirectoryOnly);
            MyViewModel.PathRows.Clear();
            var pathRow = new PathRow
            {
                Name = MyViewModel.SelectProject.Path,
                ExistIni = File.Exists(Path.Combine(MyViewModel.SelectProject.Path, SettingIni))
            };
            MyViewModel.PathRows.Add(pathRow);
            foreach (var dir in dirs)
            {
                var item = new PathRow
                {
                    Name = dir,
                    ExistIni = File.Exists(Path.Combine(dir, SettingIni))
                };
                MyViewModel.PathRows.Add(item);
            }

            MyDataGrid.ItemsSource = MyViewModel.PathRows;
        }

        private void SetButton_Click(object sender, RoutedEventArgs e)
        {
            CheckExcel();
            var button = (Button)sender;
            if (!string.IsNullOrEmpty(button.ToolTip.ToString()))
            {
                var cAutogenSetting = new SettingWindow(Path.Combine(button.ToolTip.ToString(), SettingIni))
                {
                    Left = Left + Width,
                    Top = Top
                };
                cAutogenSetting.ShowDialog();
            }

            var pathRows = new ObservableCollection<PathRow>();
            foreach (var row in MyViewModel.PathRows)
            {
                row.ExistIni = File.Exists(Path.Combine(row.Name, SettingIni));
                var newListViewRow = new PathRow
                {
                    Name = row.Name,
                    ExistIni = row.ExistIni
                };
                pathRows.Add(newListViewRow);
            }

            MyDataGrid.ItemsSource = pathRows;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            var cnt = GetType().ToString().Split('.').Length;
            var name = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1)) + ".HelpFiles.";
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            var helpWindow = new HelpWindow().Download(name, resourceNames, assembly);
            helpWindow.Left = Left + Width;
            helpWindow.Top = Top;
            helpWindow.ShowDialog();
        }

        private void CheckStatus()
        {
            Settings.Default.WatchFolder = WatchFolder.Text;
            Settings.Default.Save();
            if (!string.IsNullOrEmpty(WatchFolder.Text))
            {
                WatchingButton.IsEnabled = true;
                MyViewModel.ProjectRows.Clear();
                if (!Directory.Exists(WatchFolder.Text))
                    return;
                var dirs = Directory.GetDirectories(WatchFolder.Text, "*.*",
                    SearchOption.TopDirectoryOnly);
                if (!dirs.Any()) return;
                foreach (var dir in dirs)
                {
                    var project = Path.GetFileName(dir);
                    MyViewModel.ProjectRows.Add(new ProjectRow { Project = project, Path = dir });
                }

                MyViewModel.SelectProject = MyViewModel.ProjectRows.First();
                ProjectComboBox.SelectedIndex = 0;
            }

            ProjectComboBox.DataContext = MyViewModel;
        }

        private void SetNLog(string fileName)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                DeleteOldFileOnStartup = false,
                FileName = fileName,
                Layout = @"${date:format=yyyy-MM-dd HH\:mm\:ss} | ${level:upperCase=true} | ${message}"
            };
            config.AddTarget("file", fileTarget);
            var rule = new LoggingRule("*", LogLevel.Trace, fileTarget);
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            CheckExcel();
        }

        private void CheckExcel()
        {
            var hasExcel = HasExcel();
            if (hasExcel)
                MessageBox.Show(@"Please close all excel and Igxl !!!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            WatchingButton.IsEnabled = !HasExcel();
        }

        private void WatchFolder_OnTextChanged(object sender, EventArgs e)
        {
            CheckStatus();
        }

        #region Watch rule

        private void Button_Watching_Click(object sender, RoutedEventArgs e)
        {
            if (HasExcel())
            {
                MessageBox.Show(@"Please close all excel and Igxl !!!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                _watcher = new FileSystemWatcher(WatchFolder.Text);
                _watcher.Changed += OnChanged;
                _watcher.Created += OnCreated;
                _watcher.Renamed += OnRenamed;
                _watcher.Error += OnError;
                _watcher.Filter = "*.xls*";
                _watcher.IncludeSubdirectories = true;
                _watcher.EnableRaisingEvents = true;
                _watcher.NotifyFilter = NotifyFilters.Attributes
                                        | NotifyFilters.CreationTime
                                        | NotifyFilters.DirectoryName
                                        | NotifyFilters.FileName
                                        | NotifyFilters.LastAccess
                                        | NotifyFilters.LastWrite
                                        | NotifyFilters.Security
                                        | NotifyFilters.Size;

                int.TryParse(CheckingPeriodBox.Text, out _period);
                int.TryParse(MaxPatternWaitTime.Text, out _maxWaitTime);

                Hide();
                AddTrayIcon();
            }
        }

        private bool HasExcel()
        {
            try
            {
                var excel = Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            AddTask(e);
        }

        private void AddTask(FileSystemEventArgs e)
        {
            var dir = Path.GetDirectoryName(e.FullPath);
            if (Path.GetExtension(e.FullPath) != ".log" && dir != null)
            {
                var iniFileName = Path.Combine(Path.GetDirectoryName(dir), SettingIni);
                var queueFiles = new QueueFile
                {
                    InputFile = e.FullPath,
                    IniFile = iniFileName,
                    Time = TimeProvider.Current.Now
                };
                _queueFiles.Enqueue(queueFiles);
                if (!_worker.IsBusy)
                    _worker.RunWorkerAsync();
            }
        }

        private void OnRenamed(object sender, RenamedEventArgs e)
        {
            AddTask(e);
        }

        private void OnError(object sender, ErrorEventArgs e)
        {
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var logger = LogManager.GetCurrentClassLogger();
            while (_queueFiles.Count() != 0)
            {
                QueueFile queueFile;
                _queueFiles.TryPeek(out queueFile);

                if (queueFile != null)
                {
                    var processLog = Path.ChangeExtension(queueFile.InputFile, ".log");
                    SetNLog(processLog);

                    var inputFileChange = new InputFileChange(queueFile);
                    if (!inputFileChange.DoTask(_period, _maxWaitTime))
                    {
                        Dequeue(processLog, inputFileChange);
                    }
                    else
                    {
                        logger.Error("[" + EnumNLogMessage.Input + "] " + "Not finished and please check errors !!! ");
                        Dequeue(processLog, inputFileChange);
                    }
                    queueFile.OutputProcessLog = processLog;

                    logger.Trace("Mail to " + queueFile.MailTo + " ...");
                    new TeradyneMail().SendMail(queueFile.MailTo, queueFile);
                }
            }
        }

        private void Dequeue(string processLog, InputFileChange inputFileChange)
        {
            QueueFile result;
            _queueFiles.TryDequeue(out result);
            if (File.Exists(processLog) && !string.IsNullOrEmpty(inputFileChange.QueueFile.Output))
                if (Path.GetDirectoryName(processLog) != inputFileChange.QueueFile.Output)
                {
                    if (!Directory.Exists(inputFileChange.QueueFile.Output))
                        Directory.CreateDirectory(inputFileChange.QueueFile.Output);
                    File.Copy(processLog, Path.Combine(inputFileChange.QueueFile.Output, Path.GetFileName(processLog)),
                        true);
                }
        }

        #endregion

        private void HelpButton_Clicked(object sender, EventArgs e)
        {
            const string helpFile = @".\HelpFiles\Auto Test System Introduction_20220711.pptx";
            if (File.Exists(helpFile))
                Process.Start(helpFile);
        }

        #region Tray

        private void RemoveTrayIcon()
        {
            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }
        }

        private void AddTrayIcon()
        {
            if (_trayIcon != null) return;

            _trayIcon = new NotifyIcon();
            _trayIcon.Icon = Properties.Resources.Teradyne_T;
            _trayIcon.Text = @"Auto Test System";
            _trayIcon.Visible = true;
            _trayIcon.Click += TrayIconClick;
        }

        private void TrayIconClick(object sender, EventArgs e)
        {
            Show();
            RemoveTrayIcon();
            if (_watcher != null)
                _watcher.EnableRaisingEvents = false;
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            RemoveTrayIcon();
        }

        #endregion
    }
}