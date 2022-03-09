using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckPmicLeakage : PreCheckBase
    {
        #region Constructor

        public PreCheckPmicLeakage(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
        {
        }

        #endregion

        #region Member Function

        protected override bool CheckBusiness()
        {
            return true;
        }

        #endregion
    }
}