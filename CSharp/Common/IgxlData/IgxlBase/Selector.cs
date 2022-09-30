using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]
    public class Selector
    {
        public Selector()
        {
        }

        public Selector(string name, string value)
        {
            SelectorName = name;
            SelectorValue = value;
        }

        public string SelectorName { set; get; }
        public string SelectorValue { set; get; }
    }
}