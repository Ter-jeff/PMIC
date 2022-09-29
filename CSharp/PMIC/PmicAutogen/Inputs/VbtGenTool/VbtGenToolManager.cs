using OfficeOpenXml;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using System;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.VbtGenTool
{
    public class VbtGenToolManager
    {
        public List<TestParameterSheet> TestParameterSheets = new List<TestParameterSheet>();
        public List<VbtGenTestPlanSheet> VbtGenTestPlanSheets = new List<VbtGenTestPlanSheet>();

        #region Member Function

        public void CheckAll(List<ExcelWorkbook> workbooks)
        {
            #region Pre check

            foreach (var workbook in workbooks)
                foreach (var sheet in workbook.Worksheets)
                    if (sheet.Name.EndsWith("_TestParameter", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var testParameterReader = new TestParameterReader();
                        var testParameterSheet = testParameterReader.ReadSheet(sheet);
                        if (testParameterSheet != null && !string.IsNullOrEmpty(testParameterSheet.Block))
                        {
                            TestParameterSheets.Add(testParameterSheet);
                            break;
                        }
                    }

            foreach (var workbook in workbooks)
                foreach (var sheet in workbook.Worksheets)
                {
                    var testPlanReader = new VbtGenTestPlanSheetReader();
                    var vbtGenTestPlanSheet = testPlanReader.SheetReader(sheet);
                    if (vbtGenTestPlanSheet != null)
                        VbtGenTestPlanSheets.Add(vbtGenTestPlanSheet);
                }

            #endregion

            #region Post check

            #endregion
        }

        #endregion
    }
}