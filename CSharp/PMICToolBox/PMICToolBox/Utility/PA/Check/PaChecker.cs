using PmicAutomation.Utility.PA.Input;
using Library.Function.ErrorReport;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutomation.Utility.PA.Check
{
    public class PaChecker
    {
        public void CheckPaFile(Dictionary<string, PaSheet> paSheetDic)
        {
            foreach (KeyValuePair<string, PaSheet> sheet in paSheetDic)
            {
                foreach (PaRow row in sheet.Value.Rows)
                {
                    int count = sheet.Value.Rows.Count(p =>
                        p.BumpName.Equals(row.BumpName, StringComparison.OrdinalIgnoreCase) &&
                        p.Site.Equals(row.Site, StringComparison.OrdinalIgnoreCase) &&
                        p.PaType.Equals(row.PaType, StringComparison.OrdinalIgnoreCase));
                    if (count > 1)
                    {
                        string errMsg = "Pin name " + row.BumpName + " are duplicate !!!";
                        ErrorManager.AddError(PaErrorType.Duplicated, sheet.Value.Name, row.RowNum,
                            sheet.Value.BumpNameIndex, errMsg);
                    }

                    if (string.IsNullOrEmpty(row.InstrumentType))
                    {
                        string errMsg = string.Format("The instrument of {1} - {0} can not be recognized.",
                            row.Assignment, row.BumpName);
                        ErrorManager.AddError(PaErrorType.FormatError, sheet.Value.Name, row.RowNum,
                            sheet.Value.AssignmentIndex, errMsg);
                    }
                }
            }

            foreach (KeyValuePair<string, PaSheet> sheet in paSheetDic)
            {
                if (sheet.Value.Rows.Count == 0)
                {
                    string errMsg = @" format is incorrect, please check it again!";
                    ErrorManager.AddError(PaErrorType.FormatError, sheet.Value.Name, 1, errMsg);
                }
            }
        }
    }
}