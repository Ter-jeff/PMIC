using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AutomationCommon.EpplusErrorReport;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.VbtGenTool.Checker
{
    public class PatternChecker
    {
        public void Check(List<VbtGenTestPlanSheet> vbtGenTestPlanSheets)
        {
            try
            {
                if (vbtGenTestPlanSheets == null || !vbtGenTestPlanSheets.Any()) return;

                foreach (var sheet in vbtGenTestPlanSheets)
                {
                    var colIndex = sheet.HeaderIndex["REGISTER/MACRO NAME"];
                    foreach (var row in sheet.RowList)
                    {
                        if (!row.Command.Equals("TEST_SET_UP_PATTERN", StringComparison.OrdinalIgnoreCase)) continue;

                        if (InputFiles.PatternListMap.GetTimeSet(row.RegisterMacroName) == "TBD")
                        {
                            var errorMessage =
                                string.Format("The Pattern Name: {0} can not be found in Pattern List CSV!",
                                    row.RegisterMacroName);
                            EpplusErrorManager.AddError(PmicErrorType.MissingPattern, ErrorLevel.Error, sheet.SheetName,
                                row.RowNum, colIndex, errorMessage, row.RegisterMacroName);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }
    }
}