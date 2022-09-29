using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using System;

namespace AutoProgram.Writer
{
    public class UpdateMainFlow
    {
        public SubFlowSheet Work(SubFlowSheet mainFlow, string flowSheet)
        {
            //var index = mainFlow.FlowRows.FindIndex(x => x.Parameter
            //    .Equals("Flow_DCTEST_IDS", StringComparison.OrdinalIgnoreCase));
            var index = mainFlow.FlowRows.FindIndex(x => x.Parameter
                .Equals("Flow_DC_Conti", StringComparison.OrdinalIgnoreCase));
            var row = new FlowRow { OpCode = "call", Parameter = flowSheet };
            if (index == -1)
            {
                mainFlow.FlowRows.Add(row);
            }
            else
            {
                mainFlow.FlowRows.Insert(index + 1, row);
                for (var i = index + 2; i < mainFlow.FlowRows.Count; i++)
                {
                    var flowRow = mainFlow.FlowRows[i];
                    flowRow.OpCode = FlowRow.OpCodeNop;
                }
            }
            return mainFlow;
        }
    }
}