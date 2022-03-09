namespace IgxlData.IgxlBase
{
    public abstract class PinBase :IgxlRow
    {
        #region Property
        public string PinName { get; set; }
        public string PinType { get; set; }
        public string ChannelType { get; set; }
        public string InstrumentType { get; set; }
        #endregion

        #region Constructor
        protected PinBase(string pinName, string pinType)
        {
            PinName = pinName;
            PinType = pinType;
            ChannelType = "";
        }
        #endregion
    }
}