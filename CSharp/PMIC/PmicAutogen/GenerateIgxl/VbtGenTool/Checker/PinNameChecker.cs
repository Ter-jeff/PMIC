using CommonLib.Enum;
using CommonLib.ErrorReport;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace PmicAutogen.GenerateIgxl.VbtGenTool.Checker
{
    public class PinNameChecker
    {
        public void Check(List<VbtGenTestPlanSheet> vbtGenTestPlanSheets)
        {
            try
            {
                if (vbtGenTestPlanSheets == null || !vbtGenTestPlanSheets.Any()) return;

                if (TestProgram.IgxlWorkBk.PinMapPair.Value == null) return;

                foreach (var sheet in vbtGenTestPlanSheets)
                {
                    var colIndex = sheet.HeaderIndex["PIN"];
                    foreach (var row in sheet.RowList)
                    {
                        if (string.IsNullOrEmpty(row.Pin)) continue;

                        foreach (var pin in row.Pin.Split(','))
                            if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(pin.ToUpper()) &&
                                !TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin.ToUpper()))
                            {
                                var errorMessage = string.Format("The Pin Name: {0} can not be found in PinMap!",
                                    row.Pin);
                                ErrorManager.AddError(EnumErrorType.MissingPinName, EnumErrorLevel.Error,
                                    sheet.SheetName, row.RowNum, colIndex, errorMessage, pin);
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