using PmicAutomation.MyControls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.VbtGenToolTemplate
{
    public partial class VbtGenToolGenerator : MyForm
    {
        public VbtGenToolGenerator()
        {
            InitializeComponent();

            HelpButtonClicked += VbtGenToolGenerator_HelpButtonClicked;
        }

        private void Btn_TCM_Click(object sender, EventArgs e)
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

                new VbtGenToolTemplateMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(FileOpen_TCM.ButtonTextBox.Text))
            {
                FileOpen_OutputPath.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_TCM.ButtonTextBox.Text);
            }

            if (!File.Exists(FileOpen_TCM.ButtonTextBox.Text))
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

        private void VbtGenToolGenerator_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }
    }
}