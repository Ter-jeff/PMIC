using System.Collections.Generic;

namespace nWireDefinition.InputModel
{
    public class Field
    {
        public string FieldName { get; set; }
        public List<string> PortNames { get; set; }
        public List<string> CycleIndice { get; set; }

        public Field()
        {
            this.FieldName = string.Empty;
            this.PortNames = new List<string>();
            this.CycleIndice = new List<string>();
        }
    }
}
