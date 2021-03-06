namespace IgxlData.IgxlBase
{
    public class PatSetRow
    {
        public string ColumnA { get; set; }
        public string PatternSet { get; set; }
        public string TdGroup { get; set; }
        public string TimeDomain { get; set; }
        public string Enable { get; set; }
        public string File { get; set; }
        public string Burst { get; set; }
        public string StartLabel { get; set; }
        public string StopLabel { get; set; }
        public string Comment { get; set; }

        public void AddComment(string text)
        {
            if (string.IsNullOrEmpty(Comment))
                Comment = text;
            else
                Comment +=";"+ text;
        }
    }
}
