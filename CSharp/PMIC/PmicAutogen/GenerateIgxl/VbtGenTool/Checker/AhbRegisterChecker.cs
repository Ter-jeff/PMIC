using CommonLib.Enum;
using CommonLib.ErrorReport;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace PmicAutogen.GenerateIgxl.VbtGenTool.Checker
{
    public class AhbRegisterChecker
    {
        public void Check(List<VbtGenTestPlanSheet> vbtGenTestPlanSheets, AhbRegisterMapSheet ahbRegSheet)
        {
            try
            {
                if (vbtGenTestPlanSheets == null || !vbtGenTestPlanSheets.Any()) return;

                if (ahbRegSheet == null || ahbRegSheet.AhbRegRows == null) return;

                foreach (var sheet in vbtGenTestPlanSheets)
                {
                    var colIndex = sheet.HeaderIndex["REGISTER/MACRO NAME"];
                    foreach (var row in sheet.RowList)
                    {
                        if (!(Regex.IsMatch(row.Command, "^AHB_WRITE", RegexOptions.IgnoreCase) ||
                              Regex.IsMatch(row.Command, "^AHB_READ", RegexOptions.IgnoreCase)))
                            continue;

                        if (Regex.IsMatch(row.Command, "^AHB_WRITE_OPTION", RegexOptions.IgnoreCase))
                            continue;

                        if (!ahbRegSheet.AhbRegRows.Exists(x =>
                                x.RegName.Equals(row.RegisterMacroName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var errorMessage =
                                string.Format("The AHB Register Name: {0} can not be found in AHB Register Map!",
                                    row.RegisterMacroName);
                            ErrorManager.AddError(EnumErrorType.MissingRegister, EnumErrorLevel.Error,
                                sheet.SheetName, row.RowNum, colIndex, errorMessage, row.RegisterMacroName);
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