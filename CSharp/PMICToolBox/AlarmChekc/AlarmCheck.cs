//------------------------------------------------------------------------------
// Copyright (C) 2021 Teradyne, Inc. All rights reserved.
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
// 2022-2-14  Terry Zhang     #312       Initial creation
//------------------------------------------------------------------------------ 


using IgxlData.IgxlManager;
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

namespace AlarmChekc
{
    [Serializable]
    public partial class AlarmCheck : Form
    {
        private string _testprogram = string.Empty;
        private string _outputpath = string.Empty;

        public AlarmCheck()
        {
            InitializeComponent();
        }

        private void B_IGXLTestProgram_Click(object sender, EventArgs e)
        {
            string l_strTesterProgram = string.Empty;

            OpenFileDialog l_dlg = new OpenFileDialog();
            l_dlg.Filter = "Test Program|*.igxl";
            DialogResult l_dialogResult = l_dlg.ShowDialog();

            if (l_dialogResult == DialogResult.OK)
            {
                l_strTesterProgram = l_dlg.FileName;
                this.T_TestProgram.Text = l_strTesterProgram;
                this._testprogram = l_strTesterProgram;

                string l_strInputPath = Path.GetDirectoryName(l_strTesterProgram);
                this.T_OutputPath.Text = Path.Combine(l_strInputPath,"Output");
                this._outputpath = this.T_OutputPath.Text;
            }
            else
            {
                return;//do nothing
            }
        }

        private void B_OutputPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog l_diaglog = new FolderBrowserDialog();

            DialogResult l_Result = l_diaglog.ShowDialog();

            if (l_Result==DialogResult.OK)
            {
                string l_strOutput = l_diaglog.SelectedPath;
                this.T_OutputPath.Text = l_strOutput;
                this._outputpath = l_strOutput;
            }
            else
            {
                //do nothing
            }
        }

        private void B_Strart_Click(object sender, EventArgs e)
        {
            //validate input info
            if(!File.Exists(this._testprogram))
            {
                MessageBox.Show("Test program file is not exist!");
                return;
            }
            else
            {
                //do nothing
            }

            if(!Directory.Exists(this._outputpath))
            {
                try
                {
                    Directory.CreateDirectory(this._outputpath);
                }
                catch(Exception de)
                {
                    MessageBox.Show("Create output path failed! " + de.Message);
                }
            }
            else
            {
                //do nothing
            }

            try
            {

                IgxlProgram l_IGXLTestProgram = IGXLTestProgramParser.getIGXLTestProgram(this._testprogram);
                List<string> l_lstModules = l_IGXLTestProgram.Modules;

                BL_ModuleParser l_ModuleParser = new BL_ModuleParser(l_lstModules);
                List<ReportItem> l_lstReportItem = l_ModuleParser.getReportItem();

                BL_ReportWritter l_ReportWritter = new BL_ReportWritter(l_lstReportItem, this._outputpath);
                l_ReportWritter.ExportToFile();

                if (this.chkAlarmCheck.Checked)
                {
                    BL_AlarmSettingProcesser l_AlarmSettingProcesser = new BL_AlarmSettingProcesser(l_lstModules);
                    string l_strModuleOutput = Path.Combine(this._outputpath, "Module_AfterModify");
                    if(!Directory.Exists(l_strModuleOutput))
                    {
                        Directory.CreateDirectory(l_strModuleOutput);
                    }
                    l_AlarmSettingProcesser.GenerateNewSetting(l_strModuleOutput);
                }
                else
                {
                    //do nothing
                }

                MessageBox.Show("Alarm check finished!");

            }
            catch(Exception ex)
            {
                MessageBox.Show("Process failed! "+ex.Message);
            }
            
        }

        private void buttonTemplate_Click(object sender, EventArgs e)
        {

        }
    }
}
