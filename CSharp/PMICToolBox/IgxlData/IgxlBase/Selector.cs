using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class Selector
    {
        public string SelectorName { set; get; }
        public string SelectorValue { set; get; }

        public Selector()
        {
        }
        public Selector(string name, string value)
        {
            SelectorName = name;
            SelectorValue = value;
        }
    }
}
