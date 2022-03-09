using PmicAutomation.MyControls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.VbtGenerator
{
    public partial class VbtGeneratorFrom : MyForm
    {
        public VbtGeneratorFrom()
        {
            InitializeComponent();

            HelpButtonClicked += VbtGeneratorFrom_HelpButtonClicked;
        }

        private void Btn_Template_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.TemplateFile, true) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_Table_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.Excel) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_BasFile_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.BasFile, true) == null)
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

                new VbtGeneratorMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(FileOpen_Template.ButtonTextBox.Text))
            {
                if (FileOpen_Template.ButtonTextBox.Text.Contains(','))
                    FileOpen_OutputPath.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_Template.ButtonTextBox.Text.Split(',').First());
                else
                    FileOpen_OutputPath.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_Template.ButtonTextBox.Text);
            }

            if (!string.IsNullOrEmpty(FileOpen_Table.ButtonTextBox.Text))
            {
                FileOpen_BasFile.Enabled = false;
            }

            if (!string.IsNullOrEmpty(FileOpen_BasFile.ButtonTextBox.Text))
            {
                FileOpen_Table.Enabled = false;
            }

            if (string.IsNullOrEmpty(FileOpen_Template.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FileOpen_Table.ButtonTextBox.Text))
            {
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_Table.ButtonTextBox.Text) &&
                string.IsNullOrEmpty(FileOpen_BasFile.ButtonTextBox.Text))
            {
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text))
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

        private void VbtGeneratorFrom_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }
    }
}