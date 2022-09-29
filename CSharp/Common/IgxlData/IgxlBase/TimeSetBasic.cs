using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Name}")]
    [Serializable]
    public class TSet
    {
        public TSet()
        {
            TimingRows = new List<TimingRow>();
        }

        public string Name { get; set; }
        public string CyclePeriod { get; set; }
        public List<TimingRow> TimingRows { get; set; }

        public void AddTimingRows(List<TimingRow> timingRows)
        {
            TimingRows.AddRange(timingRows);
        }

        public void AddTimingRow(TimingRow timingRow)
        {
            TimingRows.Add(timingRow);
        }
    }
}