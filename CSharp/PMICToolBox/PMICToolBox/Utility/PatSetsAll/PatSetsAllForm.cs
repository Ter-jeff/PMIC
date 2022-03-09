using PmicAutomation.MyControls;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PatSetsAll
{
    public partial class PatSetsAllForm : MyForm
    {
        public PatSetsAllForm()
        {
            InitializeComponent();

            this.Load += PatSetsAllForm_Load;

            HelpButtonClicked += PatSetsAll_HelpButtonClicked;
        }

        private void PatSetsAllForm_Load(object sender, EventArgs e)
        {
            Array arr = Enum.GetValues(typeof(IGXLVersionEnum));
            object[] igxlveisions = new object[arr.Length];
            for (int i = 0; i < arr.Length; i++)
            {
                igxlveisions[i] = (arr.GetValue(i).ToString());
            }
            this.comboBoxIgxlVersion.Items.AddRange(igxlveisions);
            this.comboBoxIgxlVersion.SelectedIndex = 0;
        }

        private void Btn_InputPath_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(DefaultPath))
                DefaultPath = Directory.GetCurrentDirectory();
            if (PathDialog(sender,true) == null)
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

                new PatSetsAllMain(this).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {

            if (string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text) &&
                !string.IsNullOrEmpty(FileOpen_InputPath.ButtonTextBox.Text))
            {
                FileOpen_OutputPath.ButtonTextBox.Text=Directory.GetParent(FileOpen_InputPath.ButtonTextBox.Text).FullName;
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

        private void PatSetsAll_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }
    }
}