using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class ChannelMapRow : IgxlItem
    {
        #region Constructor

        public ChannelMapRow()
        {
            Sites = new List<string>();
            DeviceUnderTestPinName = "";
            DeviceUnderTestPackagePin = "";
            Type = "";
            Comment = "";
            InstrumentType = "";
        }

        #endregion

        #region Property

        public string DeviceUnderTestPinName { get; set; }
        public string DeviceUnderTestPackagePin { get; set; }
        public string Type { get; set; }
        public List<string> Sites { get; set; }
        public string Comment { get; set; }
        public string InstrumentType { get; set; }
        public int RowNum;

        #endregion
    }
}