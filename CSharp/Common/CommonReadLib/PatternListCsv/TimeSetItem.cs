namespace CommonReaderLib.PatternListCsv
{
    public class TimeSetItem
    {
        public TimeSetItem()
        {
            Version = -1;
            TimeMod = "";
        }

        public int Version { get; set; }
        public string TimeMod { get; set; }
    }
}