using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PmicAutomation.MyControls;
using System.IO;

namespace PmicAutomation.Utility.TCMID
{
    public partial class TCMIDForm : MyForm
    {
        List<string> inputFiles;

        public TCMIDForm()
        {
            InitializeComponent();
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
            richTextBox.Refresh();
        }

        public string GetTPVersion()
        {
            return tb_version.Text.Trim();
        }

        private void CheckOutputPathStatus()
        {
            if (!string.IsNullOrEmpty(outputPath.ButtonTextBox.Text))
            {
                var path = outputPath.ButtonTextBox.Text + @"\Result\";
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                outputPath.ButtonTextBox.Text = path;
            }
            else
                outputPath.ButtonTextBox.Text = string.Empty;
            CheckRunButtonStatus();
        }

        private void CheckFileStatus()
        {
            if (!string.IsNullOrEmpty(inputPath.ButtonTextBox.Text))
            {
                if (!File.Exists(inputPath.ButtonTextBox.Text))
                {
                    richTextBox.Clear();
                    AppendText(string.Format("Input file not found {0}", inputPath.ButtonTextBox.Text), Color.Red);
                    inputPath.ButtonTextBox.Text = string.Empty;
                    return;
                }
            }
            else
                inputPath.ButtonTextBox.Text = string.Empty;
            CheckRunButtonStatus();
        }

        private void CheckRunButtonStatus()
        {
            if (!string.IsNullOrEmpty(inputPath.ButtonTextBox.Text) && !string.IsNullOrEmpty(outputPath.ButtonTextBox.Text))
                Btn_RunDownload.Run.Enabled = true;
        }

        private void Btn_InputPath_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
                return;

            if (!Directory.Exists(inputPath.ButtonTextBox.Text))
            {
                AppendText("Select input path does not exist!", Color.Red);
                return;
            }
            else
            {
                inputFiles = Directory.GetFiles(inputPath.ButtonTextBox.Text, "*.txt", SearchOption.TopDirectoryOnly).ToList();
                outputPath.ButtonTextBox.Text = inputPath.ButtonTextBox.Text;
                CheckOutputPathStatus();
            }
        }

        private void Btn_OutputPath_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
                return;
            CheckOutputPathStatus();
        }

        private void Btn_Run_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox.Clear();
                new TCMIDMain(this).WorkFlow(inputFiles, tcmIdObjList:null, bCompare:false, bGenFlag:true);
                AppendText("All processes completed", Color.ForestGreen);
            }
            catch (Exception exception)
            {
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void Btn_Download_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".Template.").Show();
        }

        private void TCMIDForm_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }
    }
}
