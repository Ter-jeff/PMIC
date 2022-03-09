using System.Collections.Generic;

namespace FWFrame.nWireDefinition.InputModel
{
    public class Protocol
    {
        public string Name { get; set; }
        public List<Port> Ports { get; set; }

        public Protocol()
        {
            Name = string.Empty;
            Ports = new List<Port>();
        }
    }
}
