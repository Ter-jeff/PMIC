using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckPmicIds : PreCheckBase
    {
        public PreCheckPmicIds(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
        {
        }

        protected override bool CheckBusiness()
        {
            return true;
        }
    }
}