namespace IgxlData.IgxlBase
{
    public class Pin : PinBase
    {
        #region Field
        #endregion

        #region Property
        public string Comment { get; set; }
        #endregion

        #region Constructor

        public Pin(string pinName, string pinType, string comment = "")
            : base(pinName, pinType)
        {
            Comment = comment;
        }

        #endregion

        #region Member Function



        #endregion
    }
}