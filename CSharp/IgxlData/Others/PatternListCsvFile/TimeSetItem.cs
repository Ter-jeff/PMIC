namespace IgxlData.Others.PatternListCsvFile
{
    public class TimeSetItem
    {
        public int Version { get; set; }
        public string TimeMod { get; set; }

        public TimeSetItem()
        {
            Version = -1;
            TimeMod = "";
        }
    }
}