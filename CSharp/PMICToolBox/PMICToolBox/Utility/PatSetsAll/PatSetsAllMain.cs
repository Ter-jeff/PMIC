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
// 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz"
//
//------------------------------------------------------------------------------ 

using PmicAutomation.MyControls;
using PmicAutomation.Utility.PatSetsAll.Function;
using Library.Function.ErrorReport;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PatSetsAll
{
    public class PatSetsAllMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _inputPath;
        private readonly string _outputPath;
        private readonly bool _absolutePath;
        private readonly bool _relativePath;
        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Add Start
        private readonly bool _GzOnly;
        private readonly bool _PatOnly;
        private readonly bool _GzAndPatAll;
        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Add End
        private readonly IGXLVersionEnum _igxlVersion;

        public PatSetsAllMain(PatSetsAllForm patSetsAllForm)
        {
            _appendText = patSetsAllForm.AppendText;
            _inputPath = patSetsAllForm.FileOpen_InputPath.ButtonTextBox.Text;
            _outputPath = patSetsAllForm.FileOpen_OutputPath.ButtonTextBox.Text;
            _absolutePath = patSetsAllForm.radioButton_absolutePath.Checked;
            _relativePath = patSetsAllForm.radioButton_relativePath.Checked;
            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Add Start
            _GzOnly = patSetsAllForm.rBtnGzOnly.Checked;
            _PatOnly = patSetsAllForm.rBtnPatOnly.Checked;
            _GzAndPatAll = patSetsAllForm.rBtnAll.Checked;
            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Add End
            _igxlVersion = (IGXLVersionEnum)System.Enum.Parse(typeof(IGXLVersionEnum), patSetsAllForm.comboBoxIgxlVersion.Text);
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            GenFiles();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private void GenFiles()
        {
            GenPatSetsAll();

            GenErrorReport();
        }

        private void GenPatSetsAll()
        {
            GenPatSetsAll genPatSetsAll = new GenPatSetsAll();
            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
            //genPatSetsAll.Print(_inputPath, _outputPath, _absolutePath, _igxlVersion);
            genPatSetsAll.Print(_inputPath, _outputPath, _absolutePath, _igxlVersion, _GzOnly, _PatOnly, _GzAndPatAll);
            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End
        }

        private void GenErrorReport()
        {
            if (ErrorManager.GetErrorCount() <= 0)
                return;

            _appendText.Invoke("Starting to print error report ...", Color.Red);
            string outputFile = Path.Combine(_outputPath, "Error.xlsx");
            List<string> files = new List<string>();
            ErrorManager.GenErrorReport(outputFile, files);
        }
    }
}