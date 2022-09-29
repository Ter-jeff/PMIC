using CommonLib.Enum;
using CommonLib.ErrorReport;
using PmicAutogen.Local;
using System;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PostAction.PostCheck
{
    public class DuplicateInstanceChecker
    {
        public void WorkFlow()
        {
            var duplicateInst = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value).SelectMany(y => y.InstanceRows)
                .GroupBy(p => p.TestName).Where(p => p.Count() > 1).ToList();
            foreach (var group in duplicateInst)
            {
                var testName = group.First().TestName;
                if (!string.IsNullOrEmpty(testName))
                {
                    var errorMessage = "The instance " + testName + " in instance sheet is duplicated instance!";
                    if (!ErrorManager.GetErrors().Exists(x =>
                            x.Message.Equals(errorMessage, StringComparison.CurrentCultureIgnoreCase)))
                        ErrorManager.AddError(EnumErrorType.Duplicate, EnumErrorLevel.Warning,
                            string.Join(",", group.Select(x => x.SheetName).ToList().Distinct()), group.First().RowNum,
                            errorMessage, testName, group.Count().ToString());
                }
            }
        }
    }
}