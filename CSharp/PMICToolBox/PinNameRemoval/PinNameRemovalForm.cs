using Microsoft.WindowsAPICodePack.Dialogs;
using PinNameRemoval.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace PinNameRemoval
{
    public partial class PinNameRemovalForm : Form
    {
        Ctrl ctrl;
        List<string> fileList;

        public PinNameRemovalForm()
        {
            InitializeComponent();
            //tbInputPath.Text = tbPinmapPath.Text = @"D:\Reference\VSTemp\Pattern Test\PinMap_test.txt";
            //tbInputPath.Text = @"D:\Project\PinNameRemoval\Test";
            tbInputPath.Text = Settings.Default.InputPath;
            tbPinmapPath.Text = Settings.Default.PinmapPath;
            tbOutputPath.Text = Settings.Default.OutputPath;
            tbPinName.Text = Settings.Default.PinNameToRemove;
            ctrl = new Ctrl();

            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Text = Text.Replace("[version]", version);
        }

        private async void btnRun_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbPinName.Text))
            {
                MessageBox.Show("Please input pin names to be deleted.");
                return;
            }

            if (FindInstalledIgxl() == false)
            {
                MessageBox.Show("Please install IGXL.");
                return;
            }

            btnRun.Enabled = false;
            Ctrl.SetPinNameListToDelete(tbPinName.Text);
            try
            {
                if (string.IsNullOrEmpty(tbInputPath.Text) || string.IsNullOrEmpty(tbPinmapPath.Text) || string.IsNullOrEmpty(tbOutputPath.Text)) return;
                DateTime start = DateTime.Now;

                fileList = Ctrl.GetAllFilePath(tbInputPath.Text, "*.pat|*.pat.gz");
                if (!fileList.Any())
                {
                    MessageBox.Show("There is no *.pat file found in input folder.", "Error");
                    return;
                }
                listBox1.DataSource = fileList;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = fileList.Count;
                progressBar1.Value = 0;

                List<PatternInfo> patterns = new List<PatternInfo>();
                foreach (string filePath in fileList)
                {
                    PatternInfo pi = PatternInfo.ReadPatternInfo(filePath, tbPinmapPath.Text, tbInputPath.Text, tbOutputPath.Text);
                    patterns.Add(pi);
                }
                listBox1.DataSource = await ctrl.ProcessAsync(patterns, new Progress<int>(p => progressBar1.Value = p));
                ctrl.GenerateReport(patterns);
                Settings.Default.PinmapPath = tbPinmapPath.Text;
                Settings.Default.InputPath = tbInputPath.Text;
                Settings.Default.OutputPath = tbOutputPath.Text;
                Settings.Default.PinNameToRemove = tbPinName.Text;
                Settings.Default.Save();

                TimeSpan span = DateTime.Now.Subtract(start);
                MessageBox.Show("Time spent: " + span.ToString(), "Finish");
            }
            catch (Exception ex)
            {
                string errorMsg = ex.Message;
#if DEBUG
                errorMsg += "\n" + ex.ToString();
#endif
                MessageBox.Show(errorMsg, "Error");
            }
            finally
            {
                Ctrl.tempLog = new List<string>();
                btnRun.Enabled = true;
            }
        }

        private void btnSelectInput_Click(object sender, EventArgs e)
        {
            try
            {
                //using (var ofd = new FolderBrowserDialog())
                //{
                //    if (!string.IsNullOrEmpty(tbInputPath.Text))
                //        ofd.SelectedPath = tbInputPath.Text;
                //    DialogResult result = ofd.ShowDialog();

                //    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(ofd.SelectedPath))
                //        tbInputPath.Text = ofd.SelectedPath;
                //    else
                //        return;
                //}

                CommonOpenFileDialog fb = new CommonOpenFileDialog("Select Project Folder");
                fb.Multiselect = false;
                string path = tbInputPath.Text;
                if (string.IsNullOrEmpty(path))
                {
                    fb.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                else if (Directory.Exists(path))
                {
                    fb.InitialDirectory = path;
                }
                else
                {
                    fb.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                fb.Title = "Select input folder";
                fb.IsFolderPicker = true;
                if (fb.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    tbInputPath.Text =  fb.FileName;
                    tbOutputPath.Text = Path.Combine(tbInputPath.Text, "output");
                }
                else
                {
                    tbInputPath.Text = path;
                    tbOutputPath.Text = Path.Combine(tbInputPath.Text, "output");
                }

                if (string.IsNullOrEmpty(tbInputPath.Text) == false)
                {
                    // get the file attributes for file or directory
                    FileAttributes attr = File.GetAttributes(tbInputPath.Text);
                    //detect whether its a directory or file
                    if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                    {
                        fileList = Directory.GetFiles(tbInputPath.Text, "*.pat").ToList();
                        fileList.AddRange(Directory.GetFiles(tbInputPath.Text, "*.pat.gz").ToList());
                    }
                    else
                    {
                        fileList = new List<string>();
                        fileList.Add(tbInputPath.Text);
                    }

                    listBox1.DataSource = fileList;
                }
            }
            catch (Exception ex)
            {
                string errorMsg = ex.Message;
#if DEBUG
                errorMsg += "\n" + ex.ToString();
#endif
                MessageBox.Show(errorMsg, "Error");
            }
        }

        private void btnDeletePins_Click(object sender, EventArgs e)
        {
            //try
            //{
            btnDeletePins.Enabled = false;
            DateTime start = DateTime.Now;
            string inputPath = @"D:\Project\PinNameRemoval\Test\PP_AVSA0_C_PL00_SC_CL00_SAA_UNC_AUT_ALLFV_SI_3_1_A0_1805291402.atp";
            inputPath = tbInputPath.Text;
            string nameWoE = Path.GetFileNameWithoutExtension(inputPath);
            string outputPath = inputPath.Replace(nameWoE, nameWoE + "_delete");
            bool isScan;
            List<string> res = ctrl.DeletePins(inputPath, out isScan);
            //if (isScan) MessageBox.Show("this is scan type.");
            File.WriteAllLines(outputPath, res);

            Compile cp = new Compile(outputPath, tbPinmapPath.Text);
            nameWoE = Path.GetFileNameWithoutExtension(outputPath);
            cp.OutputPath = outputPath.Replace(nameWoE, nameWoE + "_compile");
            cp.PinmapPath = tbPinmapPath.Text;
            cp.isScanType = isScan;
            cp.Run();

            List<string> msg = new List<string>();
            msg.Add(cp.Result + "\t" + cp.FileName + "\t" + cp.OutputMsg);
            listBox1.DataSource = msg;
            TimeSpan span = DateTime.Now.Subtract(start);
            MessageBox.Show("Time spent: " + span.ToString());
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message + "\n" + ex.ToString());
            //}
            //finally
            //{
            btnDeletePins.Enabled = true;
            //}
        }

        private void btnSelectPinmap_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrEmpty(dialog.FileName))
                    tbPinmapPath.Text = dialog.FileName;
            }
        }

        private void btnSelectOutput_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog fb = new CommonOpenFileDialog("Select Project Folder");
            fb.Multiselect = false;
            string path = tbInputPath.Text;
            if (string.IsNullOrEmpty(path))
            {
                fb.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            else if (Directory.Exists(path))
            {
                fb.InitialDirectory = path;
            }
            else
            {
                fb.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            fb.Title = "Select output folder";
            fb.IsFolderPicker = true;
            if (fb.ShowDialog() == CommonFileDialogResult.Ok)
            {
                tbOutputPath.Text = fb.FileName;
            }
            else
            {
                tbOutputPath.Text = path;
            }
        }

        private static bool FindInstalledIgxl()
        {
            string igxlRoot = Environment.GetEnvironmentVariable("IGXLROOT");
            if (string.IsNullOrEmpty(igxlRoot))
                return false;
            else
                return true;
        }



    }
}
