using System.Collections.Generic;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class PatInfoData
    {
        public PatInfoData()
        {
            PatInfoList = new List<HardIpReference>();
            PatInfoErrorList = new List<HardIpReference>();
        }

        public List<HardIpReference> PatInfoList { get; set; }
        public List<HardIpReference> PatInfoErrorList { get; set; }
    }
}