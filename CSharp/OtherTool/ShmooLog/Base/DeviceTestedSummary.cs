using System.Collections.Generic;

namespace ShmooLog.Base
{
    public class DeviceTestedSummary
    {
        public int DeviceNo;

        public Dictionary<string, HashSet<string>> ShmooSetupTestedInst = new Dictionary<string, HashSet<string>>();

        public DeviceTestedSummary(int deviceNo)
        {
            DeviceNo = deviceNo;
        }

        public string DieXY { get; set; }
    }
}