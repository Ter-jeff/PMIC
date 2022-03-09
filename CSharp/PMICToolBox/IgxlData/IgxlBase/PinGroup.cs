using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PinGroup : PinBase
    {
        #region Field
        private List<Pin> _pinList;
        #endregion

        #region Property
        public List<Pin> PinList
        {
            get { return _pinList; }
            set { _pinList = value; }
        }
        #endregion

        #region Constructor
        public PinGroup(string pinGrpName, string pinType)
            : base(pinGrpName, pinType)
        {
            _pinList = new List<Pin>();
        }
        #endregion

        #region Member Function
        public void AddPin(Pin pin)
        {
            if (!_pinList.Exists(a => a.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                _pinList.Add(pin);
        }

        public void AddPin(string pin, string pinType = "")
        {
            if (!_pinList.Exists(a => a.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase)))
            {
                Pin newPin = new Pin(pin, pinType);
                _pinList.Add(newPin);
            }
        }

        public void AddPins(List<Pin> pins, string comment = "")
        {
            foreach (var pin in pins)
            {
                if (!_pinList.Exists(a => a.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                {
                    Pin newPin = new Pin(pin.PinName, pin.PinType, comment);
                    _pinList.Add(newPin);
                }
            }
        }

        public void AddPins(List<string> pins, string pinType = "")
        {
            foreach (var pin in pins)
            {
                if (!_pinList.Exists(a => a.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase)))
                {
                    Pin newPin = new Pin(pin, pinType);
                    _pinList.Add(newPin);
                }
            }
        }
        #endregion
    }
}