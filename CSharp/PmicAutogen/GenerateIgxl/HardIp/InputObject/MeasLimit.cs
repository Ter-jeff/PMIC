using System;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class MeasLimit
    {
        public MeasLimit(string jobName)
        {
            JobName = jobName;
            LoLimit = "";
            HiLimit = "";
        }

        public string JobName { get; set; }
        public string LoLimit { get; set; }
        public string HiLimit { get; set; }
        public int LoHeaderIndex { get; set; }
        public int HiHeaderIndex { get; set; }
    }
}