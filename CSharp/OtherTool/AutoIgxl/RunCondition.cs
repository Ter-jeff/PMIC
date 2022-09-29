using System;
using System.Collections.Generic;

namespace AutoIgxl
{
    public class RunCondition
    {
        public string Job { get; set; }
        public string LotId { get; set; }
        public string WaferId { get; set; }
        public string SetXy { get; set; }
        public List<string> ExecEnableWords { get; set; }
        public List<string> TotalEnableWords { get; set; }
        public string OutputLog { get; set; }
        public string FinalOutputLog { get; set; }
        public string OutputReport { get; set; }
        public bool DoAll { get; set; }
        public bool OverrideFailStop { get; set; }

        public string Tester
        {
            get { return Environment.MachineName; }
        }

        public List<string> TotalSites { get; set; }

        public List<string> ExecSites { get; set; }
    }
}