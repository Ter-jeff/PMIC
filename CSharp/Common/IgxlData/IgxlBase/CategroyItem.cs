namespace IgxlData.IgxlBase
{
    public class CategoryInSpec
    {
        #region Field

        #endregion

        #region Constructor

        public CategoryInSpec(string name, string typ, string min, string max)
        {
            Name = name;
            Typ = typ;
            Min = min;
            Max = max;
        }

        public CategoryInSpec(string name)
        {
            Name = name;
        }

        #endregion

        #region Property

        public string Name { get; set; }
        public string Typ { get; set; }
        public string Min { get; set; }
        public string Max { get; set; }

        #endregion
    }
}