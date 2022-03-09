using System;
using AutomationCommon.Utility;
using Microsoft.Office.Interop.Excel;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Inputs.TestPlan.TestPlanSetValidation;
using PmicAutogen.Local.Const;

namespace PMICAutogenAddIn
{
    public partial class ThisAddIn
    {
        #region VSTO generated code

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookOpen += WorkbookOpen;
            //Application.WorkbookActivate += WorkbookActivate;
            //Application.WorkbookNewSheet += WorkbookNewSheet;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        #region Event

        private void WorkbookOpen(Workbook wb)
        {
            //
            var isPmicTestPlan = wb.IsSheetExist(PmicConst.ProjectConfig);
            SheetStructureManager.Initialize();
            if (isPmicTestPlan)
            {
                var setValidationManager = new SetValidationManager();
                setValidationManager.SetValidations(wb);
            }
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

        #endregion
    }
}