using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class Tset : IgxlItem
    {
        #region Constructor

        public Tset()
        {
            TimingRows = new List<TimingRow>();
        }

        #endregion

        #region Member Function
        public void AddTimingRows(List<TimingRow> timingRows)
        {
            TimingRows.AddRange(timingRows);
        }

        public void AddTimingRow(TimingRow timingRow)
        {
            TimingRows.Add(timingRow);
        }

        #endregion

        #region Property

        public string Name { get; set; }
        public string CyclePeriod { get; set; }
        public List<TimingRow> TimingRows { get; set; }

        #endregion
    }
}