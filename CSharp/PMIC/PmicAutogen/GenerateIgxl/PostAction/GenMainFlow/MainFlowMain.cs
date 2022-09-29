using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PostAction.GenMainFlow
{
    public class MainFlowMain
    {
        public Dictionary<IgxlSheet, string> WorkFlow()
        {
            if (InputFiles.TestPlanWorkbook == null)
                return null;
            var mainFlow = StaticTestPlan.MainFlowSheet;
            if (mainFlow != null)
            {
                var flow = new MainFlowSheet(PmicConst.FlowMain + "_" + StaticSetting.JobMap.Values.First().First());
                flow.FlowRows = mainFlow.FlowRows;

                var igxlSheets = new Dictionary<IgxlSheet, string>();
                igxlSheets.Add(flow, FolderStructure.DirMainFlow);
                return igxlSheets;
            }

            return null;
        }
    }
}