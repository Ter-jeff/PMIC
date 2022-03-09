using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class TestPlanSequence
    {
        private int _seqIndex;

        #region Constructor

        public TestPlanSequence(int startRow, int endRow, int seqIndex)
        {
            StartRow = startRow;
            EndRow = endRow;
            _seqIndex = seqIndex;
            ForceCondition = new List<string>();
        }

        #endregion

        public int StartRow { get; set; }
        public int EndRow { get; set; }

        public int SeqIndex
        {
            set { _seqIndex = value; }
            get { return _seqIndex; }
        }

        public List<string> ForceCondition { set; get; }
    }
}