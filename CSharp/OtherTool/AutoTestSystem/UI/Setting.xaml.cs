using AutoTestSystem.Setting;
using AutoTestSystem.UI.Enable;
using AutoTestSystem.UI.Site;
using IgxlData.IgxlReader;
using Microsoft.WindowsAPICodePack.Dialogs;
using MyWpf.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace AutoTestSystem.UI
{
    public partial class SettingWindow
    {
        private const string GroupName = "CAutogen";
        private readonly ComIni _comIni = new ComIni();

        private readonly string _iniFileName;

        private SettingViewModel _settingViewModel;
        private string _initialDirectory;

        private bool _isInputFail;

        public SettingWindow(string iniFileName)
        {
            _iniFileName = iniFileName;

            InitializeComponent();
        }

        public void Save()
        {
            _isInputFail = false;

            if (File.Exists(_iniFileName))
                File.Delete(_iniFileName);

            _comIni.IniWrite(_iniFileName, "MailTo", MailTo.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "TestProgram", TestProgram.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "PatternFolder", PatternFolder.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "PatternSync", PatternSync.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "JobName", JobComboBox.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "LotId", LotId.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "WaferID", WaferId.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "SetXY", SetXY.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "EnableWords", EnableWords.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "Sites", Sites.Text, GroupName);
            _comIni.IniWrite(_iniFileName, "DoAll", DoAll.IsChecked.ToString(), GroupName);
            _comIni.IniWrite(_iniFileName, "OverrideFailStop", OverrideFailStop.IsChecked.ToString(), GroupName);
        }

        public void Load()
        {
            if (File.Exists(_iniFileName))
            {

                if (_comIni.IniKeyExists(_iniFileName, "TestProgram", GroupName))
                    TestProgram.Text = _comIni.IniRead(_iniFileName, "TestProgram", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "PatternFolder", GroupName))
                    PatternFolder.Text = _comIni.IniRead(_iniFileName, "PatternFolder", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "PatternSync", GroupName))
                    PatternSync.Text = _comIni.IniRead(_iniFileName, "PatternSync", GroupName);

                if (_comIni.IniKeyExists(_iniFileName, "MailTo", GroupName))
                    MailTo.Text = _comIni.IniRead(_iniFileName, "MailTo", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "JobName", GroupName))
                    JobComboBox.SelectedItem = _comIni.IniRead(_iniFileName, "JobName", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "LotId", GroupName))
                    LotId.Text = _comIni.IniRead(_iniFileName, "LotId", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "WaferID", GroupName))
                    WaferId.Text = _comIni.IniRead(_iniFileName, "WaferID", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "SetXY", GroupName))
                    SetXY.Text = _comIni.IniRead(_iniFileName, "SetXY", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "EnableWords", GroupName))
                    EnableWords.Text = _comIni.IniRead(_iniFileName, "EnableWords", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "Sites", GroupName))
                    Sites.Text = _comIni.IniRead(_iniFileName, "Sites", GroupName);
                if (_comIni.IniKeyExists(_iniFileName, "DoAll", GroupName))
                    DoAll.IsChecked = _comIni.IniRead(_iniFileName, "DoAll", GroupName)
                        .Equals("True", StringComparison.CurrentCultureIgnoreCase);
                if (_comIni.IniKeyExists(_iniFileName, "OverrideFailStop", GroupName))
                    OverrideFailStop.IsChecked = _comIni.IniRead(_iniFileName, "OverrideFailStop", GroupName)
                        .Equals("True", StringComparison.CurrentCultureIgnoreCase);
                AddJobs(TestProgram.Text);
            }
        }

        private void AddJobs(string testProgram)
        {
            JobComboBox.Items.Clear();
            var igxlSheetReader = new IgxlSheetReader();
            var jobs = igxlSheetReader.GetJobs(testProgram);
            if (jobs == null || jobs.Count == 0)
            {
                JobComboBox.Items.Add("CP1");
                JobComboBox.Items.Add("CP2");
                JobComboBox.Items.Add("FT1");
                JobComboBox.Items.Add("FT2");
                JobComboBox.Items.Add("QA");
            }
            else
            {
                foreach (var job in jobs) JobComboBox.Items.Add(job);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            Save();

            CheckMailTo();

            if (!_isInputFail)
                DialogResult = true;
        }

        private void CheckStatus()
        {
        }

        private void CheckMailTo()
        {
            foreach (var mailTo in MailTo.Text.Split(';'))
            {
                if (!IsValidEmail(mailTo))
                    MessageBox.Show(mailTo + " is not valid !!! ");
            }
        }

        bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();
            if (trimmedEmail.EndsWith("."))
                return false;
            try
            {
                var mailAddress = new System.Net.Mail.MailAddress(email);
                return mailAddress.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }

        private string FileSelect(object sender, EnumFileFilter filter)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = new MyWindowBase.FileFilter().GetFilter(filter),
                Multiselect = false
            };

            if (!string.IsNullOrEmpty(_initialDirectory))
                openFileDialog.InitialDirectory = _initialDirectory;
            var text = ((TextBoxButton)sender).Text;
            if (!string.IsNullOrEmpty(text))
                openFileDialog.InitialDirectory = Path.GetDirectoryName(text);

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var name = openFileDialog.FileNames.First();
                _initialDirectory = Path.GetDirectoryName(name);
                Focus();
                ((TextBoxButton)sender).Focus();
                ((TextBoxButton)sender).Text = name;
                return name;
            }

            Focus();
            ((TextBoxButton)sender).Focus();
            return null;
        }

        protected string PathSelect(object sender)
        {
            if (_initialDirectory == null)
                _initialDirectory = ((TextBoxButton)sender).Text;
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = _initialDirectory,
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                Focus();
                ((TextBoxButton)sender).Focus();
                return null;
            }

            Focus();
            ((TextBoxButton)sender).Focus();
            ((TextBoxButton)sender).Text = folderBrowserDialog.FileName;
            _initialDirectory = folderBrowserDialog.FileName;
            return folderBrowserDialog.FileName;
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            Load();
        }

        private void TestProgram_ButtonClick(object sender, EventArgs e)
        {
            var selected = EnableWords.Text.Split(',').ToList();
            if (FileSelect(sender, EnumFileFilter.Igxl) == null)
                return;

            if (!string.IsNullOrEmpty(TestProgram.Text))
            {
                var pattern = Path.Combine(Path.GetDirectoryName(TestProgram.Text), "Pattern");
                if (Directory.Exists(pattern))
                    PatternFolder.Text = pattern;
            }

            _settingViewModel = InitEnableViewModel(GetEnables(), selected);
            AddJobs(TestProgram.Text);
        }

        private List<string> GetEnables()
        {
            var igxlSheetReader = new IgxlSheetReader();
            if (File.Exists(TestProgram.Text))
                return igxlSheetReader.GetEnables(TestProgram.Text);
            MessageBox.Show("Please check if " + TestProgram.Text + " is existed !!!");
            return null;
        }

        private void PatternFolder_ButtonClick(object sender, EventArgs e)
        {
            if (PathSelect(sender) == null)
                return;

            CheckStatus();
        }

        private void PatternSync_ButtonClick(object sender, EventArgs e)
        {
            if (FileSelect(sender, EnumFileFilter.BasFile) == null)
                return;

            CheckStatus();
        }

        private void EnableWord_OnClick(object sender, RoutedEventArgs e)
        {
            var selected = EnableWords.Text.Split(',').ToList();
            _settingViewModel = InitEnableViewModel(GetEnables(), selected);

            var enablesWindow = new EnablesWindow();
            enablesWindow.ListBox.ItemsSource = _settingViewModel.EnableRows;
            enablesWindow.ShowDialog();
            if (enablesWindow.DialogResult == true)
                EnableWords.Text = string.Join(",",
                    _settingViewModel.EnableRows.Where(x => x.Select).Select(x => x.EnableWord));
        }

        private SettingViewModel InitEnableViewModel(List<string> enables, List<string> selected)
        {
            var enableViewModel = new SettingViewModel();
            foreach (var enable in enables)
            {
                var flag = selected.Contains(enable, StringComparer.CurrentCultureIgnoreCase);
                enableViewModel.EnableRows.Add(new EnableRow { EnableWord = enable, Select = flag });
            }
            return enableViewModel;
        }

        private void Site_OnClick(object sender, RoutedEventArgs e)
        {
            var selects = Sites.Text.Split(',').ToList();
            if (_settingViewModel == null)
                _settingViewModel = new SettingViewModel();

            _settingViewModel.SiteRows = InitSites(GetSites(), selects);

            var siteWindow = new SiteWindow();
            siteWindow.ListBox.ItemsSource = _settingViewModel.SiteRows;
            siteWindow.ShowDialog();
            if (siteWindow.DialogResult == true)
            {
                Sites.Text = string.Join(",", _settingViewModel.SiteRows.Where(x => x.Select).Select(x => x.Site));
            }
        }

        private ObservableCollection<SiteRow> InitSites(List<string> sites, List<string> selected)
        {
            var siteRow = new ObservableCollection<SiteRow>();
            foreach (var site in sites)
            {
                var flag = selected.Contains(site, StringComparer.CurrentCultureIgnoreCase);
                siteRow.Add(new SiteRow() { Site = site, Select = flag });
            }
            return siteRow;
        }

        private List<string> GetSites()
        {
            var igxlSheetReader = new IgxlSheetReader();
            if (File.Exists(TestProgram.Text))
                return igxlSheetReader.GetSites(TestProgram.Text);
            MessageBox.Show("Please check if " + TestProgram.Text + " is existed !!!");
            return null;
        }
    }
}