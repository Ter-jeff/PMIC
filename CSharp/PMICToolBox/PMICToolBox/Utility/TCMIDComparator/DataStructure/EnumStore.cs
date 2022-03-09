using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.TCMIDComparator.DataStructure
{
    public class EnumStore
    {
        public enum CompareStatus
        {
            ADD,
            REMOVE, // testname remove
            TCMID_REMOVE,
            MODIFY,
            NA
        }
    }
}
