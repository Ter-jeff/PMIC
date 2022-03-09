using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class Tset 
    {
        #region Property
        public string Name { get; set; }
        public string CyclePeriod { get; set; }
        public List<TimingRow> TimingRows { get; set; }
        #endregion

        #region Constructor
        public Tset()
        {
            TimingRows = new List<TimingRow>();
        }
        #endregion

        #region Member Function
        public void AddTimingRow(TimingRow timingRow)
        {
            TimingRows.Add(timingRow);
        }
        #endregion
    }
}