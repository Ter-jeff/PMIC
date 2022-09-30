using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PatternFileName}")]
    public class PatSetSubRow : IgxlRow
    {
        public string PatternFileName { get; set; }
        public string Comment { get; set; }
    }
}