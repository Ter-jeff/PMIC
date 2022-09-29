using CommonLib.Enum;
using CommonLib.WriteMessage;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using PmicAutogen.Local;
using System;
using System.IO;
using System.Linq;

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
        }

        public void WorkFlow()
        {
            try
            {
                var testParameterSheet = _testParameterSheet;

                Response.Report(string.Format("Generating Instance for " + _block + " ..."), EnumMessageLevel.General, 30);
                var instanceSheet = testParameterSheet.GenInstance();

                Response.Report(string.Format("Generating TestFlow for " + _block + " ..."), EnumMessageLevel.General, 40);
                var subFlowSheet = testParameterSheet.GenFlowSheet();

                Response.Report(string.Format("Generating Bin Table for " + _block + " ..."), EnumMessageLevel.General, 50);
                var binTableRows = testParameterSheet.GenBinTableRows(_block);
                var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                binTable.AddRows(binTableRows);

                foreach (var flowSheet in subFlowSheet.Where(flowSheet => flowSheet.FlowRows.Count > 0))
                    IgxlSheets.Add(flowSheet, Path.Combine(FolderStructure.DirModulesBlock, _block));
                IgxlSheets.Add(instanceSheet, Path.Combine(FolderStructure.DirModulesBlock, _block));
                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);
            }
            catch (Exception e)
            {
                var message = "Buck AutoGen Failed: " + e.Message;
                Response.Report(message, EnumMessageLevel.Error, 100);
            }
        }
    }
}