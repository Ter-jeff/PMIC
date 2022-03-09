using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AutomationCommon.DataStructure;
using PmicAutogen.Local.Const;
using PmicAutogen.Properties;

namespace PmicAutogen.InputPackages
{
    public class InputPackage : InputPackageBase
    {
        public bool CheckInput(WriteMessage writeMessage)
        {
            if (SelectedTestPlan.Count > 1)
            {
                writeMessage.Invoke("More than one Test Plan files were selected!", MessageLevel.Error);
                return false;
            }

            if (SelectedScgh.Count > 1)
            {
                writeMessage.Invoke("More than one SCGH files were selected!", MessageLevel.Error);
                return false;
            }

            if (SelectedPatternList.Count > 1)
            {
                writeMessage.Invoke("More than one PatternList files were selected!.", MessageLevel.Error);
                return false;
            }

            var testPlan = GetSelectedTestPlan();
            var projectName = testPlan != null ? testPlan.GetProjectName() : "";

            var selectedScghForProject = GetSelectedScgh();
            if (selectedScghForProject != null)
            {
                var scghProject = selectedScghForProject.GetProjectName();
                if (string.IsNullOrEmpty(projectName)) projectName = scghProject;
                if (!projectName.Equals(scghProject, StringComparison.OrdinalIgnoreCase))
                {
                    writeMessage.Invoke("The input files belong to different project!", MessageLevel.Error);
                    return false;
                }
            }

            var selectedPatternListForProject = GetSelectedPatternListCsv();
            if (selectedPatternListForProject != null)
            {
                var patternListProject = selectedPatternListForProject.GetProjectName();
                if (string.IsNullOrEmpty(projectName)) projectName = patternListProject;
                if (!projectName.Equals(patternListProject, StringComparison.OrdinalIgnoreCase))
                {
                    writeMessage.Invoke("The input files belong to different project!", MessageLevel.Error);
                    return false;
                }
            }

            return true;
        }

        public void SetButtonStatus(PmicMainForm pmicMainForm)
        {
            ButtonStatusClear(pmicMainForm);
            var hasTestPlan = false;
            var hasScan = false;
            var hasMbist = false;
            var hasOtpFile = false;
            var hasVbtFile = false;

            var testPlan = SelectedTestPlan;
            var vbtTestPlan = SelectedVbtTestPlan;
            var scgh = SelectedScgh;
            var otpSheetList = SelectedOtpRegisterMap;

            if (testPlan.Any())
                hasTestPlan = true;

            if (vbtTestPlan.Any())
                hasVbtFile = true;

            foreach (var otpFile in otpSheetList)
                if (Regex.IsMatch(otpFile, @".yaml", RegexOptions.IgnoreCase))
                    hasOtpFile = true;

            foreach (var sheet in testPlan.SelectMany(x => x.SheetList))
                if (Regex.IsMatch(sheet, "ahb_register_map", RegexOptions.IgnoreCase))
                    hasOtpFile = true;

            foreach (var sheet in scgh.SelectMany(x => x.SheetList))
            {
                if (sheet.ToUpper().EndsWith(PmicConst.ScghScan)) hasScan = true;
                if (sheet.ToUpper().EndsWith(PmicConst.ScghMbist)) hasMbist = true;
            }

            pmicMainForm.button_Basic.Enabled = hasTestPlan;
            pmicMainForm.button_Scan.Enabled = hasScan;
            pmicMainForm.button_Mbist.Enabled = hasMbist;
            pmicMainForm.button_Otp.Enabled = hasOtpFile;
            pmicMainForm.button_VBT.Enabled = hasVbtFile;

            pmicMainForm.button_Basic.Checked = hasTestPlan;
            pmicMainForm.button_Scan.Checked = hasScan;
            pmicMainForm.button_Mbist.Checked = hasMbist;
            pmicMainForm.button_Otp.Checked = hasOtpFile;
            pmicMainForm.button_VBT.Checked = hasVbtFile;
        }

        public List<string> OpenFileDialogWithFilter(string filterType, bool boolMultiSelect = true)
        {
            var openFileDialog = new OpenFileDialog();

            if (!string.IsNullOrEmpty(Settings.Default.InputPath) && Directory.Exists(Settings.Default.InputPath))
                openFileDialog.InitialDirectory = Settings.Default.InputPath;

            openFileDialog.Multiselect = boolMultiSelect;

            openFileDialog.Title = @"Select " + filterType;
            switch (filterType)
            {
                case "Automation Source":
                    openFileDialog.Filter =
                        @"Source Files| *Test*Plan*.xls*;*VBTPOP_Gen_*.xls*;*_pat_scg*.xlsx;*SCGH*.xls*;*pat*.csv;*Bin_Cut*.xls*;*.yaml;*.otp;*.txt";
                    break;

                case "All":
                    openFileDialog.Filter = @"All Files (*.*)|*.*";
                    break;

                case "Datalog":
                    openFileDialog.Filter = @"IG-XL Standard Output (*.txt)|*.txt|All Files (*.*)|*.*";
                    break;

                case "Excel2007":
                    openFileDialog.Filter = @"Excel 2007 (*.xls*)|*.xlsx;*.xlsm;";
                    break;

                case "AutomationLight Source":
                    openFileDialog.Filter = @"Source Files| *Test*Plan*.xls*;*SCGH*.xlsx;*pat*.csv;*.xlsm;";
                    break;

                case "Validation Source":
                    openFileDialog.Filter =
                        @"Source Files| *Test*Plan*.xls*;*.xlsm;*.igxl;*.xlsx;*_pat_scg*.xlsx;*SCGH*.xlsx;*pat*.csv;*.txt;*EFUSE*.xlsx;*Bin_Cut*.xlsx;*.yaml;*.otp;*.atp;*ExceptionList*.xlsx";
                    break;

                case "Pattern Source":
                    openFileDialog.Filter = @"PatternData Source | *.txt;*.xlsx;*.xlsm;*.csv";
                    break;

                case "C651_TW":
                    openFileDialog.Filter = @"Excel 2007 or C651 TW Csv | *.xlsx;*TW*.csv";
                    break;

                case "tCfg":
                    openFileDialog.Filter = @"T-Autogen Config | *.tCfg";
                    break;

                case "TestFlowProfile":
                    openFileDialog.Filter = @"Test Flow Profile (TestFlowProfile_*.xlsx)|TestFlowProfile_*.xlsx";
                    break;

                case "PMIC":
                    openFileDialog.Filter =
                        @"Source Files| *Test*Plan*.xls*;*VBTPOP_Gen_*.xls*;*.xlsm;*.igxl;*_pat_scg*.xlsx;*pat*.csv;*.txt;*.yaml;*.otp";
                    break;
            }

            if (openFileDialog.ShowDialog() != DialogResult.OK) return null;

            Settings.Default.InputPath = Path.GetDirectoryName(openFileDialog.FileNames.Last());
            Settings.Default.Save();
            return openFileDialog.FileNames.ToList();
        }

        private void ButtonStatusClear(PmicMainForm pmicMainForm)
        {
            pmicMainForm.button_Basic.Enabled = true;
            pmicMainForm.button_Scan.Enabled = true;
            pmicMainForm.button_Mbist.Enabled = true;
            pmicMainForm.button_Otp.Enabled = true;
            pmicMainForm.button_VBT.Enabled = true;

            pmicMainForm.button_Basic.Checked = true;
            pmicMainForm.button_Scan.Checked = true;
            pmicMainForm.button_Mbist.Checked = true;
            pmicMainForm.button_Otp.Checked = true;
            pmicMainForm.button_VBT.Checked = true;
        }
    }
}