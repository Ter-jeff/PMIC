using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckBscanChar : PreCheckBase
    {
        #region Constructor

        public PreCheckBscanChar(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
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
