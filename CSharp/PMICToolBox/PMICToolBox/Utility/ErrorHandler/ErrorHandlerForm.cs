using PmicAutomation.MyControls;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PmicAutomation.Utility.ErrorHandler
{
    public partial class ErrorHandlerForm : MyForm
    {
        public ErrorHandlerForm()
        {
            InitializeComponent();
            HelpButtonClicked += ErrorHandlerForm_HelpButtonClicked;
        }

        private void ErrorHandlerForm_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
            richTextBox.Refresh();
        }

        private void Btn_Run_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox.Clear();

                new ErrorHandlerMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("Error: " + exception.Message, Color.Red);
                //AppendText("The exception was found !!!", Color.Red);
                //AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            if (!string.IsNullOrEmpty(FileOpen_ErrorHandler.ButtonTextBox.Text) /*&&*/
                /*string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text)*/)
            {
                var path = FileOpen_ErrorHandler.ButtonTextBox.Text + @"\Output\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                FileOpen_OutputPath.ButtonTextBox.Text = path;
            }
            
            Btn_RunDownload.Run.Enabled = true;
        }

        private void Btn_ErrorHandler_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
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

        private void Btn_Download_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".Template.").Show();
        }
    }
}
