using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenFlow
{
    public class LeakageFlowRowGenerator : DcTestFlowRowGenerator
    {
        #region Constructor

        public LeakageFlowRowGenerator(string sheetName) : base(sheetName)
        {
        }

        #endregion

        #region properties

        #endregion

        #region Main Methods

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var testRows = new List<FlowRow>();
            var testRow = new FlowRow();
            testRow.Job = ""; //CreateTestJob();
            testRow.OpCode = CreateTestOpCode();
            testRow.Env = ""; // CreateTestEnv();
            testRow.Enable = GenEnable(CreateTestEnable(isCz2Only), testRow.Env);
            testRow.Parameter = CreateTestParameter();
            testRow.FailAction = CreateTestFailAction();
            testRows.Add(testRow);
            SortFlowRows(GetLimitRows(testRow), testRows);
            testRows.Add(GenBinTableRow());
            return testRows;
        }

        public override FlowRow GenBinTableRow(string voltage = "")
        {
            var binTableRow = new FlowRow();
            binTableRow.Job = ""; //CreateBinTableJob();
            binTableRow.OpCode = CreateBinTableOpCode();
            binTableRow.Env = CreateBinTableEnv();
            binTableRow.Enable = CreateBinTableEnable();
            binTableRow.Parameter = CreateBinTableParameter();
            return binTableRow;
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName);
        }

        #endregion

        #region Create flow row columns methods

        #endregion
    }
}