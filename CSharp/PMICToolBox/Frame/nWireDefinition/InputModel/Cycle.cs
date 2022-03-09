using System.Collections.Generic;

namespace FWFrame.nWireDefinition.InputModel
{
    public class Cycle
    {
        public int Repeat { get; set; }
        public List<string> Data { get; set; }
        public int VectorIndex { get; set; }
        public int CycleIndex { get; set; }

        public Cycle()
        {
            Repeat = 0;
            Data = new List<string>();
            VectorIndex = 0;
            CycleIndex = 0;
        }
    }
}
