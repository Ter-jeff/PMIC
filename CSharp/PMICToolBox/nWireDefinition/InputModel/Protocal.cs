using System.Collections.Generic;

namespace nWireDefinition.InputModel
{
    public class Protocol
    {
        public string Name { get; set; }
        public List<Port> Ports { get; set; }

        public Protocol()
        {
            this.Name = string.Empty;
            this.Ports = new List<Port>();
        }
    }
}
