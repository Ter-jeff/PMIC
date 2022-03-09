using CommonLib.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OTPFileComparison
{
    public partial class frmMain : MyFormMini
    {
        private Action<string> _downLoadEvent;
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            List<string> files = FileMultiSelect(sender, EnumFileFilter.OTPFile);
            if (files == null)
            {
                return;
            }
            foreach (var file in files)
            {
                comparisonOTPTbl1.AddNewFile(file);
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            string path = PathSelect(sender);
            if (path == null)
            {
                return;
            }

            txtOutput.Text = path;
        }

        private async void btnRun_Click(object sender, EventArgs e)
        {
            List<OTPFileInfo> OTPFileList = comparisonOTPTbl1.OTPFileInfoList;

            if (!PreCheck())
                return;
            Action<string> updateUiInfo = new Action<string>((info)=>this.Invoke(new Action(()=>SetStatusLabel(info))));
            Action<int> updateProgress = new Action<int>((value) => this.Invoke(new Action(() => SetProgressBarValue(value))));

            try
            {
                await Task.Factory.StartNew(() => new FilesCompareMain(OTPFileList, this.txtOutput.Text, updateUiInfo, updateProgress).Compare());
            }catch(Exception ex)
            {
                SetStatusLabel("Failed");
                SetProgressBarValue(0);
                MessageBox.Show("Meet Error: " + ex.ToString());
            }
            ////base.StartTime = DateTime.Now;
            ////for (int i = 0; i < 100; i++)
            ////{
            ////    SetStatusLabel("Running");
            ////    System.Threading.Thread.Sleep(10);
            ////    CalculateTimeStop();
            ////    SetProgressBarValue(i);
            ////}
        }

        private bool PreCheck()
        {
            List<OTPFileInfo> OTPFileList = comparisonOTPTbl1.OTPFileInfoList;
            if (OTPFileList.Count == 0)
            {
                MessageBox.Show("Please Load OTP Files!");
                return false;
            }
            foreach (OTPFileInfo otpFile in OTPFileList)
            {
                if (!File.Exists(otpFile.FileName)) {
                    MessageBox.Show("OTP File is not exist: " + otpFile.FileName);
                    return false;
                }
            }

            if (string.IsNullOrEmpty(this.txtOutput.Text))
            {
                MessageBox.Show("Output Folder is empty");
                return false;
            }

            if (!Directory.Exists(this.txtOutput.Text))
            {
                MessageBox.Show("Output Folder is not exist: " + this.txtOutput.Text);
                return false;
            }
            return true;

        }

        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void buttonTemplate_Click(object sender, EventArgs e)
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
