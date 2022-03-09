using System;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    internal class SinglePatChildRow : PatChildRow
    {
        public SinglePatChildRow()
        {
            IsMerged = false;
        }

        #region Property

        public TestPlanRow TpRow { get; set; }

        #endregion
    }
}