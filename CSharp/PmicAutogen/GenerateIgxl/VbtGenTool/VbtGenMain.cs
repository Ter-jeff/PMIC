using System;
using System.Linq;
using AutomationCommon.DataStructure;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenBinTable;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.VbtGenTool
{
    public class VbtGenMain : MainBase
    {
        private readonly string _block;
        private readonly TestParameterSheet _testParameterSheet;

        public VbtGenMain(TestParameterSheet testParameterSheet)
        {
            _testParameterSheet = testParameterSheet;
            _block = testParameterSheet.Block;
            if (TestProgram.IgxlWorkBk.BinTblSheets == null || !TestProgram.IgxlWorkBk.BinTblSheets.Any())
            {
                var binTableMain = new BinTableMain();
                TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinTable, binTableMain.WorkFlow());
            }
        }

        public void WorkFlow()
        {
            try
            {
                var testParameterSheet = _testParameterSheet;

                Response.Report(string.Format("Generating Instance for " + _block + " ..."), MessageLevel.General, 30);
                var instSheet = testParameterSheet.GenInstance();

                Response.Report(string.Format("Generating TestFlow for " + _block + " ..."), MessageLevel.General, 40);
                var flowSheets = testParameterSheet.GenFlow();

                Response.Report(string.Format("Generating Bin Table for " + _block + " ..."), MessageLevel.General, 50);
                var binTableRows = testParameterSheet.GenBinTableRows();

                #region AddIgxlSheet

                foreach (var flowSheet in flowSheets.Where(flowSheet => flowSheet.FlowRows.Count > 0))
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirVbt, flowSheet);

                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirVbt, instSheet);


                var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                foreach (var binTableRow in binTableRows)
                    binTable.AddRow(binTableRow);

                #endregion
            }
            catch (Exception e)
            {
                var message = "Buck AutoGen Failed: " + e.Message;
                Response.Report(message, MessageLevel.Error, 100);
            }
        }
    }
}