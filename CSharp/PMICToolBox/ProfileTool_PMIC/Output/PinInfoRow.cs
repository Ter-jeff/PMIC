using System.Collections.Generic;

namespace ProfileTool_PMIC.Output
{
    public class PinInfoRow
    {
        #region Properity
        public string InstanceName { get; set; }
        public string  Pins { get; set; }
        public List<string> PinList;
        #endregion

        #region Constructor
        public PinInfoRow()
        {
            PinList = new List<string>();
        }
        #endregion
    }
}