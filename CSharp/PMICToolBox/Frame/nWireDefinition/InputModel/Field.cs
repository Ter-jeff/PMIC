using System.Collections.Generic;

namespace FWFrame.nWireDefinition.InputModel
{
    public class Field
    {
        public string FieldName { get; set; }
        public List<string> PortNames { get; set; }
        public List<string> CycleIndice { get; set; }

        public Field()
        {
            FieldName = string.Empty;
            PortNames = new List<string>();
            CycleIndice = new List<string>();
        }
    }
}
