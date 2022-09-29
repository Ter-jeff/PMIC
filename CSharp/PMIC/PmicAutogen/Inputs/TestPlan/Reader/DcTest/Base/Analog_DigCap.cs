using System;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class AnalogDigCap
    {
        public AnalogDigCap()
        {
            Id = "DCap1";
            CoherentN = "";
            Fs = "";
            Fr = "";
            M = "";
            DigcapDiscardSamplesPerBlk = "0";
            DigcapDiscardSamplesPerAdc = "0";
            PreProcessType = "N5";
            NumAdc = "1";
            NumBlk = "1";
            AdcFullScale = "";
            CurrentBlk = "0";
            UsedVar = "";
            UsedTSet = "";
        }

        public string Id { get; set; }
        public string Fs { get; set; }
        public string CoherentN { get; set; }
        public string M { get; set; }
        public string Fr { get; set; }
        public string DigcapDiscardSamplesPerBlk { get; set; }
        public string DigcapDiscardSamplesPerAdc { get; set; }
        public string PreProcessType { get; set; }
        public string NumAdc { get; set; }
        public string NumBlk { get; set; }
        public string AdcFullScale { get; set; }
        public string CurrentBlk { get; set; }

        public string UsedVar { get; set; }
        public string UsedTSet { get; set; }
    }
}