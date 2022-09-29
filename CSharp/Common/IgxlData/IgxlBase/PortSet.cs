using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PortSet
    {
        public PortSet()
        {
            PortRows = new List<PortRow>();
        }

        public PortSet(string portName)
        {
            PortName = portName;
            PortRows = new List<PortRow>();
        }

        public void AddPortRow(PortRow portRow)
        {
            PortRows.Add(portRow);
        }

        public string PortName { get; set; }
        public List<PortRow> PortRows { get; set; }
    }
}