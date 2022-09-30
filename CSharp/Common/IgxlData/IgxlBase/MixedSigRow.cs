using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]

    public class MixedSigRow : IgxlRow
    {
        public string Name { get; set; }
        public string Subset { get; set; }
        public string Type { get; set; }
        public string Id { get; set; }
        public string Fs { get; set; }
        public string N { get; set; }
        public string Fr { get; set; }
        public string M { get; set; }
        public string Usr { get; set; }
        public string Data { get; set; }
        public string Definition { get; set; }
        public string Filter { get; set; }
        public string Settings { get; set; }
        public string WaveName { get; set; }
        public string Amplitude { get; set; }
        public string Offset { get; set; }
        public string OldInstData { get; set; }
        public string Comment { get; set; }
    }
}