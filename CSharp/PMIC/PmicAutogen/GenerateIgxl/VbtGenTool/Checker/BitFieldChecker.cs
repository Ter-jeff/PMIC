using CommonLib.Enum;
using CommonLib.ErrorReport;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace PmicAutogen.GenerateIgxl.VbtGenTool.Checker
{
    public class BitFieldChecker
    {
        public void Check(List<VbtGenTestPlanSheet> vbtGenTestPlanSheets, AhbRegisterMapSheet ahbRegSheet)
        {
            try
            {
                if (vbtGenTestPlanSheets == null || !vbtGenTestPlanSheets.Any()) return;

                if (ahbRegSheet == null || ahbRegSheet.AhbRegRows == null) return;

                foreach (var sheet in vbtGenTestPlanSheets)
                {
                    var colIndex = sheet.HeaderIndex["BITFIELD NAME"];
                    foreach (var row in sheet.RowList)
                    {
                        if (string.IsNullOrEmpty(row.BitfieldName)) continue;

                        if (!ahbRegSheet.AhbRegRows.Exists(x =>
                                x.RegName.Equals(row.BitfieldName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var errorMessage =
                                string.Format("The bitField Name: {0} can not be found in AHB Register Map!",
                                    row.BitfieldName);
                            ErrorManager.AddError(EnumErrorType.MissingRegister, EnumErrorLevel.Error,
                                sheet.SheetName, row.RowNum, colIndex, errorMessage, row.BitfieldName);
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