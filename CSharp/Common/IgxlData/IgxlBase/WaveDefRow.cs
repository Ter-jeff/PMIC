﻿using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{WaveDefName}")]
    public class WaveDefRow : IgxlRow
    {
        public string WaveDefName { get; set; }
        public string WaveDefType { get; set; }
        public string WaveDefComponent { get; set; }
        public string RepeatCount { get; set; }
        public string RelativePeriod { get; set; }
        public string RelativeAmplitude { get; set; }
        public string RelativeOffset { get; set; }
        public string PrimitiveParameters { get; set; }
        public string Comment { get; set; }
    }
}