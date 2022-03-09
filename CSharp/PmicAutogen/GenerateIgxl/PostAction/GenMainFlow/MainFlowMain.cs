using System.Linq;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.PostAction.GenMainFlow
{
    public class MainFlowMain
    {
        public void WorkFlow()
        {
            if (InputFiles.TestPlanWorkbook == null) return;
            var mainFlow = StaticTestPlan.MainFlowSheet;
            if (mainFlow != null)
            {
                var flow = new SubFlowSheet(PmicConst.FlowMain + "_" + StaticSetting.JobMap.Values.First().First());
                flow.FlowRows = mainFlow.FlowRows;
                TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);
            }
        }
    }
}