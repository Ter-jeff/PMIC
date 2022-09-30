using CommonLib.Extension;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class PinGroup : IgxlRow
    {
        public PinGroup(string pinGrpName)
        {
            PinName = pinGrpName;
            PinList = new List<Pin>();
        }

        public string PinName { get; set; }

        public List<Pin> PinList { get; set; }

        private string _pinType;

        public string PinType
        {
            get
            {
                if (!string.IsNullOrEmpty(_pinType))
                    return _pinType;
                return PinList.First().PinType;
            }
            set { _pinType = value; }
        }

        public void AddPin(Pin pin)
        {
            if (!PinList.Exists(a => a.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                PinList.Add(pin);
        }

        public void AddPin(string pin, string pinType = "")
        {
            if (!PinList.Exists(a => a.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase)))
            {
                var newPin = new Pin(pin, pinType);
                PinList.Add(newPin);
            }
        }

        public void AddPins(List<Pin> pins, string comment = "")
        {
            foreach (var pin in pins)
                if (!PinList.Exists(a => a.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                {
                    var newPin = pin.DeepClone();
                    newPin.Comment = comment;
                    PinList.Add(newPin);
                }
        }

        public void AddPins(List<string> pins, string pinType = "")
        {
            foreach (var pin in pins)
                if (!PinList.Exists(a => a.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase)))
                {
                    var newPin = new Pin(pin, pinType);
                    PinList.Add(newPin);
                }
        }
    }
}