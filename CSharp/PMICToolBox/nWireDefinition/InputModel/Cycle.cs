using System.Collections.Generic;

namespace nWireDefinition.InputModel
{
    public class Cycle
    {
        public int Repeat { get; set; }
        public List<string> Data { get; set; }
        public int VectorIndex { get; set; }
        public int CycleIndex { get; set; }

        public Cycle()
        {
            this.Repeat = 0;
            this.Data = new List<string>();
            this.VectorIndex = 0;
            this.CycleIndex = 0;
        }
    }
}
