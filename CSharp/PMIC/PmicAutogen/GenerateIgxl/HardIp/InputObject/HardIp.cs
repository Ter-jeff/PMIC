using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    public class HardIp
    {
        public HardIp()
        {
            PlanSheets = new List<TestPlanSheet>();
        }

        public List<TestPlanSheet> PlanSheets { get; set; }
    }
}