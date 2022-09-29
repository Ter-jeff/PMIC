using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    internal class MergedPatChildRow : PatChildRow
    {
        public MergedPatChildRow()
        {
            IsMerged = true;
            TpRows = new List<TestPlanRow>();
        }

        #region Property

        public List<TestPlanRow> TpRows { get; set; }

        #endregion
    }
}