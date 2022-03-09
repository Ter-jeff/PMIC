using PmicAutomation.MyControls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PmicAutomation.Utility.AHBEnum
{
    public partial class AhbEnum : MyForm
    {
        public AhbEnum()
        {
            InitializeComponent();
            HelpButtonClicked += AhbEnum_HelpButtonClicked;
        }

        private void Btn_AhbRegister_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.Excel) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_OutputPath_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender,true) == null)
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
                
                new AhbMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(FileOpen_AhbRegister.ButtonTextBox.Text))
            {
                FileOpen_OutputPath.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_AhbRegister.ButtonTextBox.Text);
            }

            if (!File.Exists(FileOpen_AhbRegister.ButtonTextBox.Text))
            {
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

        private void AhbEnum_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }
    }
}