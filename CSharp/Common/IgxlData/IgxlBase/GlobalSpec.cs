namespace IgxlData.IgxlBase
{
    public class GlobalSpec : Spec
    {
        public string Job { get; set; }

        public GlobalSpec()
        {
        }

        public GlobalSpec(string symbol) : base(symbol)
        {
            Job = "";
        }

        public GlobalSpec(string symbol, string value = "", string job = "", string comment = "")
            : base(symbol, value, comment)
        {
            Job = job;
        }
    }
}