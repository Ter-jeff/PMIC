using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Basic.GenPatSet
{
    public class SubPatInfo
    {
        public SubPatInfo()
        {
            Subroutine = new List<string>();
            VmVector = null;
        }

        public List<string> Subroutine { get; set; }
        public string VmVector { get; set; }
    }
}