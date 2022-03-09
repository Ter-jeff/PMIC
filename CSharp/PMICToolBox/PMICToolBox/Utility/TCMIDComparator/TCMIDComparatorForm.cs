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
using PmicAutomation.Utility.TCMID;
using PmicAutomation.Utility.TCMID.Business;
using PmicAutomation.Utility.TCMIDComparator.Business;

namespace PmicAutomation.Utility.TCMIDComparator
{
    public partial class TCMIDComparatorForm : MyForm
    {
        public TCMIDComparatorForm()
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

        private bool CheckFileStatus()
        {
            if (!string.IsNullOrEmpty(inputFile1.ButtonTextBox.Text) && !string.IsNullOrEmpty(inputFile2.ButtonTextBox.Text))
            {
                if (!File.Exists(inputFile1.ButtonTextBox.Text))
                {
                    richTextBox.Clear();
                    AppendText(string.Format("Input file 1 not found {0}", inputFile1.ButtonTextBox.Text), Color.Red);
                    inputFile1.ButtonTextBox.Text = string.Empty;
                    return false;
                }
                if (!File.Exists(inputFile2.ButtonTextBox.Text))
                {
                    richTextBox.Clear();
                    AppendText(string.Format("Input file 2 not found {0}", inputFile1.ButtonTextBox.Text), Color.Red);
                    inputFile2.ButtonTextBox.Text = string.Empty;
                    return false;
                }
                if (inputFile1.ButtonTextBox.Text.Equals(inputFile2.ButtonTextBox.Text))
                {
                    richTextBox.Clear();
                    AppendText("Input files are the same one", Color.Red);
                    return false;
                }
                return true;
            }

            return false;
        }

        private void CheckRunButtonStatus()
        {
            if (CheckFileStatus())
                Btn_RunDownload.Run.Enabled = true;
            else
                Btn_RunDownload.Run.Enabled = false;
        }

        private void Btn_InputFile1_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.BasFile, false) == null)
                return;

            if (!File.Exists(inputFile1.ButtonTextBox.Text))
            {
                AppendText("Select input file 1 does not exist!", Color.Red);
                return;
            }
        }

        private void Btn_InputFile2_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.BasFile, false) == null)
                return;

            if (!File.Exists(inputFile2.ButtonTextBox.Text))
            {
                AppendText("Select input file 2 does not exist!", Color.Red);
                return;
            }
            else
            {
                outputPath.ButtonTextBox.Text = Path.GetDirectoryName(inputFile2.ButtonTextBox.Text);
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
                List<string> inputList = new List<string>();
                inputList.Add(inputFile1.ButtonTextBox.Text);
                inputList.Add(inputFile2.ButtonTextBox.Text);

                richTextBox.Clear();

                List<TcmIDGenBase> tcmIdObjList = new List<TcmIDGenBase>();
                new TCMIDMain(this).WorkFlow(inputList, tcmIdObjList, bCompare:true, bGenFlag:false);
                new TcmIDCompare(this, tcmIdObjList).Process();

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
