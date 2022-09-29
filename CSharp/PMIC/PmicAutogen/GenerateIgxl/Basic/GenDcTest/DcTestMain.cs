using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest;
using PmicAutogen.Local;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Basic.GenDcTest
{
    public class DcTestMain
    {
        public DcTestMain(TestPlanSheet dcTestSheet)
        {
            DcTestSheet = dcTestSheet;
        }

        private TestPlanSheet DcTestSheet { get; }

        internal Dictionary<IgxlSheet, string> Workflow()
        {
            var instanceSheet = DcTestSheet.GenInsSheet("TestInst_" + DcTestSheet.SheetName);

            var subFlowSheet = DcTestSheet.GenFlowSheet("Flow_" + DcTestSheet.SheetName);

            var binTableRows = DcTestSheet.GenBinTableRows();
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);

            var igxlSheets = new Dictionary<IgxlSheet, string>();
            igxlSheets.Add(instanceSheet, FolderStructure.DirDcTestFunc);
            igxlSheets.Add(subFlowSheet, FolderStructure.DirDcTestFunc);
            return igxlSheets;
        }
    }
}