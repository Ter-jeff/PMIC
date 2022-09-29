using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.GenDcTest.Writer;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest
{
    public class TestPlanSheet
    {
        private const string PrefixDctest = "DCTEST_";

        public TestPlanSheet()
        {
            PatternRows = new List<PatternRow>();
            PatternItems = new List<HardIpPattern>();
            MultipleInit = false;
            PlanHeaderIdx = new Dictionary<string, int>();
        }

        public string SheetName { get; set; }
        public List<PatternRow> PatternRows { get; set; }
        public List<HardIpPattern> PatternItems { get; set; }
        public string ForceStr { get; set; }
        public int ForceIndex { get; set; }
        public int MeasIndex { get; set; }
        public bool MultipleInit { get; set; }
        public Dictionary<string, int> PlanHeaderIdx { get; set; }

        public string Block
        {
            get { return SheetName.ToUpper().Replace(PrefixDctest, "").Replace(" ", "").Replace("_", ""); }
        }

        internal BinTableRows GenBinTableRows()
        {
            var blockBinGenerator = new DcTestBinTableWriter(SheetName, PatternItems);
            return blockBinGenerator.GenBinTableRows();
        }

        internal SubFlowSheet GenFlowSheet(string flowSheetName)
        {
            var flowSheetGenerator = new DcTestFlowWriter(SheetName, PatternItems);
            return flowSheetGenerator.GenFlowSheet(flowSheetName);
        }

        internal InstanceSheet GenInsSheet(string insSheetName)
        {
            var blockInstanceGenerator = new DcTestInstanceWriter(SheetName, PatternItems);
            return blockInstanceGenerator.GenInsSheet(insSheetName);
        }
    }
}