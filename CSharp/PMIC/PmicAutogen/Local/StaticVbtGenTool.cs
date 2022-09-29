using PmicAutogen.Inputs.VbtGenTool;
using PmicAutogen.Inputs.VbtGenTool.Reader;
using System.Collections.Generic;

namespace PmicAutogen.Local
{
    public static class StaticVbtGenTool
    {
        public static List<TestParameterSheet> TestParameterSheets;
        public static List<VbtGenTestPlanSheet> VbtGenTestPlanSheets;

        public static void AddSheets(VbtGenToolManager vbtGenToolManager)
        {
            VbtGenTestPlanSheets = vbtGenToolManager.VbtGenTestPlanSheets;
            TestParameterSheets = vbtGenToolManager.TestParameterSheets;
        }
    }
}