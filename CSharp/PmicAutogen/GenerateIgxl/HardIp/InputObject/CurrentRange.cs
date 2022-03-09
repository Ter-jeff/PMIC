using System;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class CurrentRange
    {
        public CurrentRange()
        {
            JobName = "";
            Value = "";
        }

        public CurrentRange(string jobName, string value)
        {
            JobName = jobName;
            Value = value;
        }

        public string JobName { get; set; }
        public string Value { get; set; }
    }
}