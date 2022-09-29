namespace PmicAutogen.GenerateIgxl.Basic.GenConti.Base
{
    public class DcTestLimit
    {
        public DcTestLimit(string header, string hiLimitShort, string loLimitShort, string hiLimitOpen,
            string loLimitOpen)
        {
            Header = header;
            HiLimitShort = hiLimitShort;
            LoLimitShort = loLimitShort;
            HiLimitOpen = hiLimitOpen;
            LoLimitOpen = loLimitOpen;
        }

        public string Header { set; get; }
        public string HiLimitShort { set; get; }
        public string LoLimitShort { set; get; }
        public string HiLimitOpen { set; get; }
        public string LoLimitOpen { set; get; }
    }
}