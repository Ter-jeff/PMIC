using CommonLib.Enum;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using DebugPlanChecker.Properties;
using MyWpf.Controls;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace DebugPlanChecker
{
    public partial class DebugPlanCheckerMainWindow
    {
        public DebugPlanCheckerMainWindow()
        {
            HelpButtonClick += HelpButton_Clicked;

            InitializeComponent();

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            Title += " " + version;

            LoadSettings();
            Response.Progress = WriteMessage();
        }

        private void LoadSettings()
        {
            DebugPlan.Text = Settings.Default.DebugTestPlan;
            PatternPath.Text = Settings.Default.PatthernFolder;
            OutputFolder.Text = Settings.Default.OutputFolder;
            TestProgram.Text = Settings.Default.TestProgram;
            OutputFolder.Text = Settings.Default.OutputFolder;
            CheckStatus();
        }

        public Progress<ProgressStatus> WriteMessage()
        {
            var progress = new Progress<ProgressStatus>();
            progress.ProgressChanged += (o, info) =>
            {
                Run r = new Run(info.Message);
                if (info.Level == EnumMessageLevel.Error)
                    r.Foreground = new SolidColorBrush(Colors.Red);
                else
                    r.Foreground = new SolidColorBrush(Colors.Black);
                Paragraph paragraph = new Paragraph(r);
                MyRichTextBox.Document.Blocks.Add(paragraph);
                MyRichTextBox.ScrollToEnd();
            };
            return progress;
        }

        private void OutputFolder_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void PatternFolder_Click(object sender, EventArgs e)
        {
            PathSelect(sender);
            CheckStatus();
        }

        private void DebugPlan_Click(object sender, EventArgs e)
        {
            if (FileSelect(sender, EnumFileFilter.Excel) == null)
                return;
            CheckStatus();
        }

        private void TestProgram_Click(object sender, EventArgs e)
        {
            if (FileSelect(sender, EnumFileFilter.Igxl) == null)
                return;
            CheckStatus();
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(DebugPlan.Text) &&
                !string.IsNullOrEmpty(OutputFolder.Text))
            {
                ButRun.IsEnabled = true;
            }
            Settings.Default.DebugTestPlan = DebugPlan.Text;
            Settings.Default.PatthernFolder = PatternPath.Text;
            Settings.Default.TestProgram = TestProgram.Text;
            Settings.Default.OutputFolder = OutputFolder.Text;
            Settings.Default.Save();
        }

        private async void Run_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Cursor = Cursors.Wait;
                ((Control)sender).IsEnabled = false;
                StartTime = TimeProvider.Current.Now;
                MyRichTextBox.Document.Blocks.Clear();
                var debugPlan = DebugPlan.Text;
                var patternPath = PatternPath.Text;
                var outputFolder = OutputFolder.Text;
                var testProgram = TestProgram.Text;
                await Task.Factory.StartNew(() => new DebugPlanCheckerMain(debugPlan, patternPath, testProgram,
                     outputFolder).WorkFlow());

            }
            catch { }
            finally
            {
                ((Control)sender).IsEnabled = true;
                MyStatusBar.StatusText = string.Format("Process time : {0}", (TimeProvider.Current.Now - StartTime).ToString(@"hh\:mm\:ss"));
                Cursor = null;
            }
        }

        private void HelpButton_Clicked(object sender, EventArgs e)
        {
            const string helpFile = @".\HelpFiles\Debug_Plan_Checker_Introduction_20220922.pptx";
            if (File.Exists(helpFile))
                Process.Start(helpFile);
        }
    }
}
