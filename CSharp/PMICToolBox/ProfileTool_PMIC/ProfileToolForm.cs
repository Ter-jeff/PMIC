using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using CommonLib.Controls;
using IgxlData.IgxlBase;
using IgxlData.IgxlManager;
using IgxlData.IgxlReader;

namespace ProfileTool_PMIC
{
    public partial class ProfileToolForm : MyForm
    {
        private string _tempPath;

        public ProfileToolForm()
        {
            InitializeComponent();
            EnvironmentPrecheck();
            HelpButtonClicked += HelpButton_Clicked;
        }

        private void EnvironmentPrecheck()
        {
            string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT");
            if (!Directory.Exists(oasisRootFolder))
            {
                MessageBox.Show("Oasis not found.\nTest Program Modify is disabled.", "Oasis not found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Button1.Enabled = false;
            }
        }

        #region Tab1
        private void button_run1_Click(object sender, EventArgs e)
        {
            Button1.Enabled = false;
            CalculateTimeStart();

            var exportMain = new IgxlManagerMain();
            exportMain.ExportWorkBook(FileOpen_TestProgram.ButtonTextBox.Text, _tempPath);
            AppendText("Exporting test program to txt ...", Color.Blue);

            var profileToolMain = new ProfileGeneraterMain(this, _tempPath);
            profileToolMain.WorkFlow();

            CalculateTimeStop();
            Button1.Enabled = true;
        }


        private void buttonLoad_CorePowerList_Click(object sender, EventArgs e)
        {
            Reset();
            if (FileSelect(sender, EnumFileFilter.Txt) == null)
                return;

            CheckStatus();
        }

        private void buttonLoad_ExecutionProfile_Click(object sender, EventArgs e)
        {
            Reset();
            if (FileSelect(sender, EnumFileFilter.Txt) == null)
                return;

            CheckStatus();
        }

        private void buttonLoad_TestProgarm_Click(object sender, EventArgs e)
        {
            Reset();
            if (FileSelect(sender, EnumFileFilter.TestProgram) == null)
                return;

            if (string.IsNullOrEmpty(_tempPath))
                _tempPath = Path.Combine(Path.GetDirectoryName(FileOpen_TestProgram.ButtonTextBox.Text), "Temp");

            var exportMain = new IgxlManagerMain();
            exportMain.ExportWorkBook(FileOpen_TestProgram.ButtonTextBox.Text, _tempPath);
            AppendText("Exporting test program to txt ...", Color.Blue);

            GetchannelMap();

            CheckStatus();
        }

        private void GetchannelMap()
        {
            AppendText("Parsing channel map, and Please wait ...", Color.Blue);
            var readChanMapSheet = new ReadChanMapSheet();
            var channelMapSheets = readChanMapSheet.GetIgxlSheets(_tempPath, SheetType.DTChanMap);
            if (channelMapSheets.Any())
            {
                ComboBox_ChanMap.Items.AddRange(channelMapSheets.Select(x => x.Name).ToArray());
                ComboBox_ChanMap.SelectedIndex = 0;
                AppendText("Please select the channel map ...", Color.Blue);
            }
            else
                AppendText("Please check if there are any channel map in the test program ...", Color.Red);
        }

        private void buttonLoad_Output1_Click(object sender, EventArgs e)
        {
            if (PathSelect(sender) == null)
                return;

            CheckStatus();
        }

        private void CheckStatus()
        {


            if (!string.IsNullOrEmpty(FileOpen_TestProgram.ButtonTextBox.Text))
            {
                if (string.IsNullOrEmpty(FileOpen_OutputPath1.ButtonTextBox.Text))
                    FileOpen_OutputPath1.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_TestProgram.ButtonTextBox.Text);
            }

            if (!string.IsNullOrEmpty(FileOpen_ExecutionProfile.ButtonTextBox.Text))
            {
                if (string.IsNullOrEmpty(FileOpen_OutputPath1.ButtonTextBox.Text))
                    FileOpen_OutputPath1.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_ExecutionProfile.ButtonTextBox.Text);
            }

            if (Directory.Exists(FileOpen_OutputPath1.ButtonTextBox.Text) &&
                (File.Exists(FileOpen_ExecutionProfile.ButtonTextBox.Text) &&
                File.Exists(FileOpen_TestProgram.ButtonTextBox.Text) &&
                ComboBox_ChanMap.Items.Count > 0))
                Button1.Enabled = true;
            else
                Button1.Enabled = false;
        }

        private void HelpButton_Clicked(Object sender, EventArgs e)
        {
            var cnt = GetType().ToString().Split('.').Length;
            var name = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1)) + ".HelpFiles.";
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            new MyDownloadForm().Download(name, resourceNames, assembly).Show();
        }
        #endregion

        #region Tab2
        private void buttonLoad_profile_Click(object sender, EventArgs e)
        {
            if (PathSelect(sender) == null)
                return;

            CheckStatus2();
        }

        private void buttonLoad_Output2_Click(object sender, EventArgs e)
        {
            if (PathSelect(sender) == null)
                return;

            CheckStatus2();
        }

        private void CheckStatus2()
        {
            if (radioButtonIndividual.Checked)
                groupBoxFilter.Enabled = true;
            else
                groupBoxFilter.Enabled = false;

            if (!string.IsNullOrEmpty(FileOpen_ProfilePath1.ButtonTextBox.Text))
            {
                if (string.IsNullOrEmpty(FileOpen_OutputPath2.ButtonTextBox.Text))
                    FileOpen_OutputPath2.ButtonTextBox.Text = Path.GetDirectoryName(FileOpen_ProfilePath1.ButtonTextBox.Text);
            }


            double value;
            if (!double.TryParse(textBoxPulseWidth.Text, out value))
            {
                Button2.Enabled = false;
                AppendText("The filter width is not a number !", Color.Red);
            }
            else
            {
                if (Directory.Exists(FileOpen_ProfilePath1.ButtonTextBox.Text) && Directory.Exists(FileOpen_OutputPath2.ButtonTextBox.Text))
                    Button2.Enabled = true;
                else
                    Button2.Enabled = false;
            }
        }

        private void button_run2_Click(object sender, EventArgs e)
        {
            Button2.Enabled = false;
            Reset();
            CalculateTimeStart();
            var outputFiles = new List<string>();
            var outputFileName = Path.Combine(FileOpen_OutputPath2.ButtonTextBox.Text, "ProfileChart");
            outputFiles.Add(outputFileName + ".xlsx");
            outputFiles.Add(outputFileName + ".pptx");
            foreach (var outputFile in outputFiles)
            {
                if (File.Exists(outputFile))
                    if (!IsOpened(outputFile))
                        File.Delete(outputFile);
                    else
                    {
                        AppendText(string.Format("Please close file {0}...", outputFile), Color.Red);
                        Button2.Enabled = true;
                        return;
                    }
            }

            var profileAnalysis = new ProfileAnalysisMain(this, outputFileName);
            profileAnalysis.WorkFlow();

            CalculateTimeStop();
            Button2.Enabled = true;
        }

        private void textBoxChartCount_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !Char.IsDigit(e.KeyChar) && !Char.IsControl(e.KeyChar);
        }

        private void textBoxChartCount_KeyPressDouble(object sender, KeyEventArgs e)
        {
            CheckStatus2();
        }
        #endregion

        private void radioButtonManual_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonManual.Checked)
                textBoxChartCount.Enabled = true;
        }

        private void radioButtonIndividual_CheckedChanged(object sender, EventArgs e)
        {
            CheckStatus2();
        }

        private void checkBox_Power_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Power.Checked)
            {
                FileOpen_ProfilePath2.Enabled = true;
            }
            else
            {
                FileOpen_ProfilePath2.Enabled = false;
            }
        }
    }
}
