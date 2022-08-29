using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PortSet : IgxlItem
    {
        #region Constructor

        public PortSet(string portName)
        {
            PortName = portName;
            PortRows = new List<PortRow>();
        }

        #endregion

        #region Member Function

        public void AddPortRow(PortRow portRow)
        {
            PortRows.Add(portRow);
        }

        #endregion

        #region Property

        public string PortName { get; set; }
        public List<PortRow> PortRows { get; set; }

        #endregion
    }
}