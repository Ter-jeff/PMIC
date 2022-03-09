using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenConti.Base
{
    public class RelayStatus
    {
        public RelayStatus()
        {
            OpenRelayList = new List<string>();
            OffRelayList = new List<string>();
        }

        public List<string> OpenRelayList { set; get; }
        public List<string> OffRelayList { set; get; }

        public bool IsEqualStatus(RelayStatus targetStatus)
        {
            if (!OpenRelayList.Any() && !targetStatus.OpenRelayList.Any()) return true;
            if (OpenRelayList.All(p => targetStatus.OpenRelayList.Contains(p)) &&
                targetStatus.OpenRelayList.All(p => OpenRelayList.Contains(p)))
                return true;
            return false;
        }
    }
}