using PmicAutomation.MyControls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PmicAutomation.Utility.Relay
{
    public partial class Relay : MyForm
    {
        public Relay()
        {
            InitializeComponent();

            HelpButtonClicked += Relay_HelpButtonClicked;

            MessageBox.Show("Relay function might not cover all trace as Component pin report format haven't aligned between different project", 
                "Relay Tool Tips", MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void Btn_ComPin_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.Excel) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_RelayConfig_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.Excel) == null)
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

        private void Btn_Run_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox.Clear();
                StartTime = DateTime.Now;
                new RelayMain(this).WorkFlow();
                CalculateTime();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(FileOpen_ComPin.ButtonTextBox.Text))
            {
                FileOpen_OutputPath.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_ComPin.ButtonTextBox.Text);
            }

            if (!File.Exists(FileOpen_ComPin.ButtonTextBox.Text))
            {
                return;
            }

            if (!chkboxAdg1414.Checked && !File.Exists(FileOpen_RelayConfig.ButtonTextBox.Text))
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

        private void Relay_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }

        private void chkboxAdg1414_CheckedChanged(object sender, EventArgs e)
        {
            FileOpen_RelayConfig.Enabled = !chkboxAdg1414.Checked;
            CheckStatus();
        }
    }
}