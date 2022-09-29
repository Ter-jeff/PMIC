using CommonLib.Enum;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Local.Const;
using PmicAutogen.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.InputPackages
{
    public class InputPackage : InputPackageBase
    {
        public bool HasTestPlan
        {
            get { return SelectedTestPlan.Any(); }
        }

        public bool HasVbtFile
        {
            get { return SelectedVbtTestPlan.Any(); }
        }

        public bool HasOtpFile
        {
            get
            {
                foreach (var otpFile in SelectedOtpRegisterMap)
                    if (Regex.IsMatch(otpFile, @".yaml", RegexOptions.IgnoreCase))
                        return true;

                foreach (var sheet in SelectedTestPlan.SelectMany(x => x.SheetList))
                    if (Regex.IsMatch(sheet, "ahb_register_map", RegexOptions.IgnoreCase))
                        return true;

                return false;
            }
        }

        public bool HasScan
        {
            get
            {
                foreach (var sheet in SelectedScgh.SelectMany(x => x.SheetList))
                    if (sheet.ToUpper().EndsWith(PmicConst.ScghScan))
                        return true;
                return false;
            }
        }

        public bool HasMbist
        {
            get
            {
                foreach (var sheet in SelectedScgh.SelectMany(x => x.SheetList))
                    if (sheet.ToUpper().EndsWith(PmicConst.ScghMbist))
                        return true;
                return false;
            }
        }

        public bool CheckInput(WriteMessage writeMessage)
        {
            if (SelectedTestPlan.Count > 1)
            {
                writeMessage.Invoke("More than one Test Plan files were selected!", EnumMessageLevel.Error);
                return false;
            }

            if (SelectedScgh.Count > 1)
            {
                writeMessage.Invoke("More than one SCGH files were selected!", EnumMessageLevel.Error);
                return false;
            }

            if (SelectedPatternList.Count > 1)
            {
                writeMessage.Invoke("More than one PatternList files were selected!.", EnumMessageLevel.Error);
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
                    writeMessage.Invoke("The input files belong to different project!", EnumMessageLevel.Error);
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
                    writeMessage.Invoke("The input files belong to different project!", EnumMessageLevel.Error);
                    return false;
                }
            }

            return true;
        }

        public void Clear(Workbook workbook)
        {
            if (workbook != null)
                InputFiles = InputFiles.FindAll(p => p is InputTestPlan && p.Selected);
            else
                InputFiles = new List<Base.Input>();
        }

        //public void SetButtonStatus(PmicMainForm pmicMainForm)
        //{
        //    ButtonStatusClear(pmicMainForm);

        //    pmicMainForm.button_Basic.Enabled = HasTestPlan;
        //    pmicMainForm.button_Scan.Enabled = HasScan;
        //    pmicMainForm.button_Mbist.Enabled = HasMbist;
        //    pmicMainForm.button_Otp.Enabled = HasOtpFile;
        //    pmicMainForm.button_VBT.Enabled = HasVbtFile;

        //    pmicMainForm.button_Basic.Checked = HasTestPlan;
        //    pmicMainForm.button_Scan.Checked = HasScan;
        //    pmicMainForm.button_Mbist.Checked = HasMbist;
        //    pmicMainForm.button_Otp.Checked = HasOtpFile;
        //    pmicMainForm.button_VBT.Checked = HasVbtFile;
        //}

        //public void UpdateAutomationBlockStatus()
        //{
        //    BlockStatus.UpdateAutomationBlockStatus(HasTestPlan, HasScan, HasMbist, HasOtpFile, HasVbtFile);
        //}

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

            if (!(bool)openFileDialog.ShowDialog()) return null;

            Settings.Default.InputPath = Path.GetDirectoryName(openFileDialog.FileNames.Last());
            Settings.Default.Save();
            return openFileDialog.FileNames.ToList();
        }

        //private void ButtonStatusClear(PmicMainForm pmicMainForm)
        //{
        //    pmicMainForm.button_Basic.Enabled = true;
        //    pmicMainForm.button_Scan.Enabled = true;
        //    pmicMainForm.button_Mbist.Enabled = true;
        //    pmicMainForm.button_Otp.Enabled = true;
        //    pmicMainForm.button_VBT.Enabled = true;

        //    pmicMainForm.button_Basic.Checked = true;
        //    pmicMainForm.button_Scan.Checked = true;
        //    pmicMainForm.button_Mbist.Checked = true;
        //    pmicMainForm.button_Otp.Checked = true;
        //    pmicMainForm.button_VBT.Checked = true;
        //}
    }
}