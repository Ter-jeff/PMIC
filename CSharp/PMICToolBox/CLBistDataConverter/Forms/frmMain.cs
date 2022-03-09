using CLBistDataConverter.DataStructures;
using CLBistDataConverter.Libs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace CLBistDataConverter
{
    public partial class frmMain : Form
    {
        #region preprocess field
        private BackgroundWorker _BgWorker = new BackgroundWorker();
        private Action<string> _downLoadEvent;
        #endregion

        #region preprocess methods
        #region events
        private void frmMain_Load(object sender, EventArgs e)
        {
            PreInitialize();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            rTxtBoxLog.Clear();
            _BgWorker.RunWorkerAsync();
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void releaseNotesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmReleaseNote frm = new frmReleaseNote();
            frm.ShowDialog();
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout frm = new frmAbout();
            frm.ShowDialog();
        }
        #endregion
        #region  methods
        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            WorkerMain(worker, e);
        }
        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            WorkerComplete(worker, e);
        }
        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.Invoke(new Action(() => { toolStripProgressBar1.Value = e.ProgressPercentage; }));
        }
        private void PreInitialize()
        {
            this.Text = Assembly.GetExecutingAssembly().GetName().Name + " V" + Assembly.GetExecutingAssembly().GetName().Version;

            GlobalSpecs.initialize();

            #region initialize Reporter
            ReportLib.HandllerWriteMsg = new Action<string, MessageLevel>((msg, level) =>
            {
                this.Invoke(new Action(() =>
                {
                    msg = ComLib.LogTimeStemp() + " : " + msg;
                    ComLib.WriteLog(msg);

                    if (level == MessageLevel.err)
                    {
                        rTxtBoxLog.Select(rTxtBoxLog.Text.Length, 0);
                        rTxtBoxLog.Focus();
                        rTxtBoxLog.SelectionColor = Color.Tomato;
                    }
                    else if (level == MessageLevel.warn)
                    {
                        rTxtBoxLog.Select(rTxtBoxLog.Text.Length, 0);
                        rTxtBoxLog.Focus();
                        rTxtBoxLog.SelectionColor = Color.YellowGreen;
                    }
                    else
                    {
                        rTxtBoxLog.Select(rTxtBoxLog.Text.Length, 0);
                        rTxtBoxLog.Focus();
                        rTxtBoxLog.SelectionColor = Color.Black;
                    }
                    rTxtBoxLog.AppendText(msg + Environment.NewLine);
                    rTxtBoxLog.ScrollToCaret();
                }));
            });
            ReportLib.HandllerReportPrgress = new Action<int, int>((val, max) =>
            {
                this.Invoke(new Action(() =>
                {
                    toolStripProgressBar1.Value = val;
                    toolStripProgressBar1.Maximum = max;
                }));
            });
            ReportLib.HandllerReportPrgressAndMsg = new Action<string, int, int>((msg, val, max) =>
            {
                this.Invoke(new Action(() =>
                {
                    msg = ComLib.TimeStemp() + " : " + msg;
                    ComLib.WriteLog(msg);
                    rTxtBoxLog.AppendText(msg + Environment.NewLine);
                    rTxtBoxLog.ScrollToCaret();
                    toolStripProgressBar1.Value = val;
                    toolStripProgressBar1.Maximum = max;
                }));
            });
            ReportLib.HandllerReportState = new Action<string>((msg) =>
            {
                this.Invoke(new Action(() =>
                {
                    toolStripStatusLabelState.Text = msg;
                }));
            });
            #endregion

            #region initialize backgroundworker
            _BgWorker.DoWork += bgWorker_DoWork;
            _BgWorker.RunWorkerCompleted += bgWorker_RunWorkerCompleted;
            _BgWorker.ProgressChanged += bgWorker_ProgressChanged;
            _BgWorker.WorkerReportsProgress = true;
            #endregion
        }
        #endregion
        #endregion

        public frmMain()
        {
            InitializeComponent();
        }

        private void btnInput_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdialog = new OpenFileDialog();
            ofdialog.Title = "Select DataLog File";
            ofdialog.Filter = "DataLog File|*.txt";
            if (ofdialog.ShowDialog() == DialogResult.OK)
            {
                txtDatalogFilePath.Text = ofdialog.FileName;
                //ReportLib.ReportState("Txt File");
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdialog = new FolderBrowserDialog();
            if (fbdialog.ShowDialog() == DialogResult.OK)
            {
                txtOutputFolder.Text = fbdialog.SelectedPath;
            }
        }

        private void WorkerMain(BackgroundWorker worker, DoWorkEventArgs e)
        {
            try
            {
                ReportLib.ReportProgress(0);
                ReportLib.WriteMsg("Start");

                ReportLib.WriteMsg("Do Pre Check.");
                if (!PreCheck())
                {
                    return;
                }

                ReportLib.ReportProgress(10);
                ReportLib.WriteMsg("Read DataLog File : " + txtDatalogFilePath.Text);
                DatalogReader fr = new DatalogReader();
                List<CLBistDie> clBistData = fr.Read(txtDatalogFilePath.Text);
                ReportLib.ReportProgress(60);
                ReportLib.WriteMsg("Edit CLBist Data");
                CLBistEditer editer = new CLBistEditer();
                editer.Edit(clBistData);
                ReportLib.ReportProgress(80);
                FileInfo fi = new FileInfo(txtDatalogFilePath.Text);
                string newFileName = Path.Combine(txtOutputFolder.Text, fi.Name.Replace(fi.Extension, "") + "_" + ComLib.TimeStemp() + ".csv");

                ReportLib.WriteMsg("Write File : " + newFileName);
                FilerWriter fw = new FilerWriter();
                fw.Write(clBistData, newFileName);
                ReportLib.ReportProgress(100);
                ReportLib.WriteMsg("Done");
            }catch(Exception ex)
            {
                ReportLib.WriteMsg("Error: " + ex.Message.ToString(),MessageLevel.err);
                ReportLib.ReportProgress(0);
            }
        }

        private void WorkerComplete(BackgroundWorker worker, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                toolStripStatusLabelMsg.Text = "Error";
            }
            else if (e.Cancelled)
            {
                toolStripStatusLabelMsg.Text = "Canceled";
            }
            else
            {
                toolStripStatusLabelMsg.Text = "Done";
            }

            //ReportLib.ReportState("");
            btnRun.Enabled = true;
        }

        private bool PreCheck()
        {
            if (txtDatalogFilePath.Text.Trim() == "")
            {
                ReportLib.WriteMsg("DataLog File Is Empty", MessageLevel.err);
                return false;
            }

            if (!File.Exists(txtDatalogFilePath.Text.Trim()))
            {
                ReportLib.WriteMsg("DataLog File Not Exists", MessageLevel.err);
                return false;
            }

            if (txtOutputFolder.Text.Trim() == "")
            {
                ReportLib.WriteMsg("Output Folder Is Empty", MessageLevel.err);
                return false;
            }

            if (!Directory.Exists(txtOutputFolder.Text.Trim()))
            {
                ReportLib.WriteMsg("Output Folder Not Exists", MessageLevel.err);
                return false;
            }

            return true;
        }

        private void button_download_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            _downLoadEvent(resourceName);
        }

        public void SetDownLoadEvent(Action<string> inputEvent)
        {
            _downLoadEvent = inputEvent;
        }
    }
}
