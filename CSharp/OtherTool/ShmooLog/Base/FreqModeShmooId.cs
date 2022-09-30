using System.Collections.Generic;

namespace ShmooLog.Base
{
    public class FreqModeShmooId
    {
        public string DieXY = "NONE";
        public Dictionary<string, double> FreqHvccDict = new Dictionary<string, double>();

        public Dictionary<string, double> FreqLvccDict = new Dictionary<string, double>();
        public string LotId = "NONE";
        public string Site = "0"; //反正只是列印用 不用宣告int
    }
}