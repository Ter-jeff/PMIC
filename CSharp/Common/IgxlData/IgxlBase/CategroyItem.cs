using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]
    public class CategoryInSpec
    {
        public CategoryInSpec(string name, string typ, string min, string max)
        {
            Name = name;
            Typ = typ;
            Min = min;
            Max = max;
        }

        public CategoryInSpec(string name)
        {
            Name = name;
        }
        public string Name { get; set; }
        public string Typ { get; set; }
        public string Min { get; set; }
        public string Max { get; set; }
    }
}