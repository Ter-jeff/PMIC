namespace CommonLib.FormatCheck
{
    public class SheetConfig
    {
        public string SheetName { get; set; }
        public string FirstHeaderName { get; set; }
        public string HeaderName { get; set; }
        public bool Optional { get; set; }
        public EnumColumn Type { get; set; }                 //for dynamic

        public SheetConfig()
        {
            SheetName = "";
            FirstHeaderName = "";
            HeaderName = "";
            Optional = false;
            Type = EnumColumn.None;
        }
    }
}