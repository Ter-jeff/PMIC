using System;
using System.Drawing;
using System.Linq;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.OTPRegisterMap
{
    public partial class OtpRegisterMapFrom : MyForm
    {
        public OtpRegisterMapFrom()
        {
            InitializeComponent();
            HelpButtonClicked += OtpRegisterMap_HelpButtonClicked;
        }

        private void Btn_OtpFilesFile_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.OtpFile, true) == null)
            {
                return;
            }
            CheckStatus();
        }

        private void Btn_YamlFile_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.YamlFile) == null)
            {
                return;
            }
            CheckStatus();
        }

        private void Btn_OutputPath_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_RegMap_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.BasFile) == null)
            {
                return;
            }
            CheckStatus();
        }

        private void Btn_Run_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox.Clear();

                new OtpRegisterMapMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text))
            {
                Btn_RunDownload.Run.Enabled = false;
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_Yaml.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FilesOpen_Otp.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FileOpen_RegMap.ButtonTextBox.Text))
            {
                Btn_RunDownload.Run.Enabled = false;
                return;
            }

            if (!string.IsNullOrEmpty(FileOpen_Yaml.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FilesOpen_Otp.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FileOpen_RegMap.ButtonTextBox.Text))
            {
                Btn_RunDownload.Run.Enabled = false;
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_Yaml.ButtonTextBox.Text) &&
                !string.IsNullOrEmpty(FilesOpen_Otp.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FileOpen_RegMap.ButtonTextBox.Text))
            {
                Btn_RunDownload.Run.Enabled = false;
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_Yaml.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FilesOpen_Otp.ButtonTextBox.Text) &&
                !string.IsNullOrEmpty(FileOpen_RegMap.ButtonTextBox.Text))
            {
                Btn_RunDownload.Run.Enabled = false;
                return;
            }

            Btn_RunDownload.Run.Enabled = true;
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
            richTextBox.Refresh();
        }

        private void Btn_Download_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".Template.").Show();
        }

        private void OtpRegisterMap_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }        
    }
}