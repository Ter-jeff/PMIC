using PmicAutomation.Utility.Relay.Input;
using System.Collections.Generic;

namespace PmicAutomation.Utility.Relay.Base
{
    public class AdgMatrix
    {
        public List<ComPinRow> DevicePins = new List<ComPinRow>();
        public string Name;

        public List<ComPinRow> ResourcePins = new List<ComPinRow>();
    }
}