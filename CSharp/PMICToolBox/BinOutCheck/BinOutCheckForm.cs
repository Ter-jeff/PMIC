//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut
//------------------------------------------------------------------------------ 
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
using OfficeOpenXml;

namespace BinOutCheck
{
    public partial class BinOutCheckForm : Form
    {
        private Action<string> _downLoadEvent;
        public BinOutCheckForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //var tpPath = @"D:\Kimi\PMIC Validation\MP12P\ND34_CP1_X10_E00_201223_V02A_CORR_SVN4860.igxl";
            //var exportfolder = Path.Combine(Directory.GetCurrentDirectory(), "tmp", "exportProg");
            //TestProgramUtility.ExportWorkBookCmd(tpPath, exportfolder);
            //var igxlProgram = new IgxlProgram(tpPath);
            //igxlProgram.LoadIgxlProgramAsync(@"D:\Kimi\PMIC Validation\BinOut Check\ND34_CP1_X10_E00_201223_V02A_CORR_SVN4860");

            //var main = new BinoutCheckMain(igxlProgram);
            //main.WorkFlow();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Multiselect = false,
                Title = "Open IG-XL File",
                Filter = "IG-XL File(*.xlsm;*.igxl)|*.xlsm;*.igxl"
            };

            if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames == null) return;
            txtTestProgram.Text = ofd.FileName;
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            //var exportfolder = Path.Combine(Directory.GetCurrentDirectory(), "tmp", "exportProg");
            //if (!Directory.Exists(exportfolder))
            //    Directory.CreateDirectory(exportfolder);
            //else
            //{
            //    Directory.Delete(exportfolder, true);
            //    Directory.CreateDirectory(exportfolder);
            //}

            //await Task.Factory.StartNew(() => TestProgramUtility.ExportWorkBookCmd(txtTestProgram.Text, exportfolder));

            //var igxlProgram = new IgxlProgram(txtTestProgram.Text);
            //await Task.Factory.StartNew(() => igxlProgram.LoadIgxlProgramAsync(exportfolder));
            try
            {
                // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
                //var main = new BinoutCheckMain(txtTestProgram.Text);
                GuiArgs guiArgs = new GuiArgs();
                guiArgs.TestPlanPath = txtTestProgram.Text;
                guiArgs.DataLogPath = txtDataLog.Text;
                var main = new BinoutCheckMain(guiArgs);
                // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
                await Task.Factory.StartNew(main.WorkFlow);
                MessageBox.Show(@"Done");
            }catch (Exception ex)
            {
                MessageBox.Show("Check BinOut Error: " + ex.ToString());
            }

            btnStart.Enabled = true;
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

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
        private void btnSelectDataLog_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Multiselect = false,
                Title = "Open Data Log File",
                Filter = "Data Log File(*.txt)|*.txt"
            };

            if (ofd.ShowDialog() != DialogResult.OK || ofd.FileNames == null) return;
            txtDataLog.Text = ofd.FileName;
        }
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end
    }
}
