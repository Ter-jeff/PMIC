using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckPmicLeakage : PreCheckBase
    {
        public PreCheckPmicLeakage(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
        {
        }

        protected override bool CheckBusiness()
        {
            return true;
        }
    }
}