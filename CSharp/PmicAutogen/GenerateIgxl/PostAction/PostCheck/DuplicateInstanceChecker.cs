using System;
using System.Linq;
using AutomationCommon.EpplusErrorReport;
using PmicAutogen.Local;

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
                    if (!EpplusErrorManager.GetErrors().Exists(x =>
                        x.Message.Equals(errorMessage, StringComparison.CurrentCultureIgnoreCase)))
                        EpplusErrorManager.AddError(DuplicateInstance.Duplicate.ToString(), ErrorLevel.Warning,
                            string.Join(",", group.Select(x => x.SheetName).ToList().Distinct()), group.First().RowNum,
                            errorMessage, testName, group.Count().ToString());
                }
            }
        }
    }
}