using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Symbol}")]
    public abstract class Spec : IgxlRow
    {
        public string Symbol { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }

        protected Spec()
        {
        }
        protected Spec(string symbol)
        {
            Symbol = symbol;
            Value = "";
            Comment = "";
        }

        protected Spec(string symbol, string value = "", string comment = "")
        {
            Symbol = symbol;
            Value = value;
            Comment = comment;
        }
    }
}