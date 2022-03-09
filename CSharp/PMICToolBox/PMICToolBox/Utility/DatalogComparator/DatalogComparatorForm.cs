using Library;
using Library.Common;
using Microsoft.WindowsAPICodePack.Dialogs;
using PmicAutomation.MyControls;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PmicAutomation.Utility.DatalogComparator
{
    public partial class DatalogComparatorForm:MyForm
    {
        private BackgroundWorker _bgWorkerForHardWorkCheck = null;
        private CommonData data = CommonData.GetInstance();

        public DatalogComparatorForm()
        {
            InitializeComponent();
        }

        private void FileOpen_Output_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
            {
                return;
            }

            CheckStatus();
        }

        protected string FileDialog(object sender, bool isFolderPicker, bool multiSelect = false)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = DefaultPath;
            if (dialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                return null;
            }
            DefaultPath = dialog.FileNames.First();
            var fileNames = multiSelect ?
                string.Join(",", dialog.FileNames) :
                dialog.FileNames.First();
            ((Control)sender).Parent.Text = fileNames;
            return fileNames;
        }

        protected string PathDialog(object sender, bool isFolderPicker, bool multiSelect = false)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = DefaultPath;
            dialog.IsFolderPicker = isFolderPicker;
            dialog.Multiselect = multiSelect;
            if (dialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                return null;
            }
            DefaultPath = dialog.FileNames.First();
            var fileNames = multiSelect ?
                string.Join(",", dialog.FileNames) :
                dialog.FileNames.First();
            ((Control)sender).Parent.Text = fileNames;
            return fileNames;
        }

        private void CheckStatus()
        {
            //if (!string.IsNullOrEmpty(FileOpen_ErrorHandler.ButtonTextBox.Text) &&
            //    string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text))
            //{
            //    var path = FileOpen_ErrorHandler.ButtonTextBox.Text + @"\Output\";
            //    if (!Directory.Exists(path))
            //    {
            //        Directory.CreateDirectory(path);
            //    }

            //    FileOpen_OutputPath.ButtonTextBox.Text = path;
            //}

            //Btn_RunDownload.Run.Enabled = true;
        }

        private void FileOpen_BaseLog_ButtonTextBoxButtonClick(object sender, EventArgs e)
        {
            if (FileDialog(sender, true) == null)
            {
                return;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                CommonData.GetInstance().OutputPath = FileOpen_Output.ButtonTextBox.Text;
                CommonData.GetInstance().BaseTxtDatalogPath = FileOpen_BaseLog.ButtonTextBox.Text;
                CommonData.GetInstance().CompareTxtDatalogPath = FileOpen_CompareLog.ButtonTextBox.Text;

                if (!Directory.Exists(FileOpen_Output.ButtonTextBox.Text))
                {
                    MessageBox.Show("Output Folder is not exist!", "Error", MessageBoxButtons.OK);
                    return;
                }

                _bgWorkerForHardWorkCheck = new BackgroundWorker();
                _bgWorkerForHardWorkCheck.WorkerReportsProgress = true;
                _bgWorkerForHardWorkCheck.DoWork += new DoWorkEventHandler(RunTestProgram);
                _bgWorkerForHardWorkCheck.ProgressChanged += new ProgressChangedEventHandler(ProgressChanged);
                _bgWorkerForHardWorkCheck.RunWorkerCompleted += new RunWorkerCompletedEventHandler(WorkerCompleted);
                _bgWorkerForHardWorkCheck.RunWorkerAsync();                
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void RunTestProgram(object sender, EventArgs e)
        {

            CommonData.GetInstance().worker = _bgWorkerForHardWorkCheck;
            MainLogic.Instance().MainFlow();
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            CommonData.GetInstance().ProgressValue = (e.ProgressPercentage.ToString());
        }

        private void WorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            CommonData.GetInstance().UIEnabled = true;
            this.Cursor = null;
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
        }

        private void buttonTemplate_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".Template.").Show();
        }
    }
}
