namespace IgxlData.IgxlBase
{
    public class LevelRow : IgxlRow
    {
        #region Property
        public int RowNum { get; set; }
        public string PinName { get; set; }
        public string Seq { get; set; }
        public string Parameter { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }
        #endregion

        #region Constructor

        public LevelRow(string pinName, string parameter, string value, string comment, int rowNum = 0)
        {
            PinName = pinName;
            Parameter = parameter;
            Value = value;
            Comment = comment;
            RowNum = rowNum;
        }

        #endregion

        #region Member Function
        public bool IsBlankRow()
        {
            if (Comment == "" &&
                Parameter == "" &&
                PinName == "" &&
                Value == "")
            {
                return true;
            }
            return false;
        }
        #endregion
    }
}