namespace IgxlData.Others.PatternListCsvFile
{
    public class PatternListCsvRow
    {
        public string PatternName { get; set; }
        public int RowNum { get; set; }
        public string LatestVersion { get; set; }
        public string Use { get; set; }
        public string Org { get; set; }
        public string TypeSpec { get; set; }
        public string TimeSetVersion { get; set; }
        public string ActualTimeSetVersion { get; set; }
        public string FileVersion { get; set; }
        public string OpCode { get; set; }
        public string ScanMode { get; set; }
        public string Halt { get; set; }
        public string OriginalTimingMode { get; set; }
        public string Check { get; set; }
        public string TpCategory { get; set; }
        public string CheckComment { get; set; }

        public PatternListCsvRow()
        {
            PatternName = "";
            RowNum = 1;
            LatestVersion = "";
            Use = "";
            Org = "";
            TypeSpec = "";
            TimeSetVersion = "";
            ActualTimeSetVersion = "";
            FileVersion = "";
            OpCode = "";
            ScanMode = "";
            Halt = "";
            OriginalTimingMode = "";
            Check = "";
            TpCategory = "";
            CheckComment = "";
        }        
    }
}
