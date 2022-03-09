using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class Pin
    {
        #region Property
        public string PinName { get; set; }
        public string PinType { get; set; }
        public string ChannelType { get; set; }
        public string InstrumentType { get; set; }
        public string Comment { get; set; }
        #endregion

        #region Constructor
        public Pin(string pinName, string pinType, string comment = "")
        {
            PinName = pinName;
            PinType = pinType;
            ChannelType = "";
            Comment = comment;
        }
        #endregion
    }
}