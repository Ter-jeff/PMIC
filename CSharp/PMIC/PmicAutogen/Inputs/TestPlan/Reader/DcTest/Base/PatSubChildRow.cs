using System;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
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