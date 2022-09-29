using System.Collections.Generic;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    public class SortLimitPin
    {
        private List<SortLimitPin> _useLimitPins;
        public string PinName { get; set; }
        public MeasPin MeasPinData { get; set; }

        public List<SortLimitPin> UseLimitPins
        {
            set { _useLimitPins = value; }
            get { return _useLimitPins ?? (_useLimitPins = new List<SortLimitPin>()); }
        }

        public void AddData(string pinName, MeasPin measPin)
        {
            UseLimitPins.Add(new SortLimitPin { PinName = pinName, MeasPinData = measPin });
        }
    }
}