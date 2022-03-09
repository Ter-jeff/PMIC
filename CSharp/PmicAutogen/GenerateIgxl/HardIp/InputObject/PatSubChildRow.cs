using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    internal class PatSubChildRow : PatChildRow
    {
        public PatSubChildRow()
        {
            TpRows = new List<TestPlanRow>();
        }

        #region Property

        public List<TestPlanRow> TpRows { get; set; }

        #endregion
    }
}