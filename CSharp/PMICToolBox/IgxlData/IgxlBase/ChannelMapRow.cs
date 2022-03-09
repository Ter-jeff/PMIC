using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class ChannelMapRow : IgxlRow
    {
        #region Property
        public string DiviceUnderTestPinName { get; set; }
        public string DiviceUnderTestPackagePin { get; set; }
        public string Type { get; set; }
        public List<string> Sites { get; set; }
        public string Comment { get; set; }

        public string InstrumentType { get; set; }
        #endregion

        #region Constructor
        public ChannelMapRow()
        {
            Sites = new List<string>();
            DiviceUnderTestPinName = "";
            DiviceUnderTestPackagePin = "";
            Type = "";
            Comment = "";
            InstrumentType = "";
        }
        #endregion
    }
}
