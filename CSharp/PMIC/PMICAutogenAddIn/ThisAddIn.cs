using CommonLib.Extension;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using PmicAutogen.Inputs.TestPlan.TestPlanSetValidation;
using PmicAutogen.Local.Const;
using PMICAutogenAddIn.Hook;
using PMICAutogenAddIn.Properties;
using System;
using System.Windows.Forms;

namespace PMICAutogenAddIn
{
    public partial class ThisAddIn
    {
        public string currentSheet = "";
        public static string lastSheet = "";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            #region go back
            Application.WorkbookOpen += WorkbookOpen;
            Application.SheetActivate += SheetActivate;
            EnableShortCuts(true);
            #endregion
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            EnableShortCuts(false);
        }

        internal static void EnableShortCuts(bool enable)
        {
            if (enable)
            {
                KbHook.SetHook();
                Globals.ThisAddIn.Application.OnKey("+^{" + Settings.Default.GoBackShortCut + "}", "");
            }
            else
            {
                KbHook.ReleaseHook();
                //Globals.ThisAddIn.Application.OnKey("+^{" + Settings.Default.GoBackShortCut + "}", "");
            }
        }

        public static void GoBack()
        {
            if (!string.IsNullOrEmpty(lastSheet))
                Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[lastSheet].Select();
        }

        private void GoLastSheet(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            GoBack();
        }

        private void SheetActivate(object Sh)
        {
            lastSheet = currentSheet;
            currentSheet = ((Worksheet)Sh).Name;
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        private void WorkbookOpen(Workbook wb)
        {
            //var isPmicTestPlan = wb.IsSheetExist(PmicConst.ProjectConfig);
            //SheetStructureManager.Initialize();
            //if (isPmicTestPlan)
            //{
            //    var setValidationManager = new SetValidationManager();
            //    setValidationManager.SetValidations(wb);
            //}

            lastSheet = Globals.ThisAddIn.Application.ActiveSheet.Name;
            currentSheet = Globals.ThisAddIn.Application.ActiveSheet.Name;

#if DEBUG
            try
            {
                if (Environment.MachineName == "DX4G433")
                    Globals.Ribbons.MyRibbon.RibbonUI.ActivateTab("tab_Autogen");
            }
            catch
            {
            }
#endif
        }

        private void SheetChange(object sh, Range range)
        {
            Workbook wb = range.Parent.Parent;
            var isPmicTestPlan = wb.IsSheetExist(PmicConst.ProjectConfig);
            if (isPmicTestPlan)
            {
                var setValidationManager = new SetValidationManager();
                setValidationManager.SetValidations(wb, true);
            }
        }
    }
}