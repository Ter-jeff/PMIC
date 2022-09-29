using System;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class AnalogSrc
    {
        public AnalogSrc()
        {
            Id = "Usrc1";
            Pin = "";
            Fs = "";
            Fi = "";
            N = "";
            M = "";
            Amp = "1";
            Vcm = "0.5";
            VcmEnable = "1";
            Bw = "80M";
            DiffMode = "1";
            UseWave = "SineWave";
        }

        public string Id { get; set; }
        public string Pin { get; set; }
        public string Fs { get; set; }
        public string Fi { get; set; }
        public string N { get; set; }
        public string M { get; set; }
        public string Amp { get; set; }
        public string Vcm { get; set; }
        public string VcmEnable { get; set; }
        public string Bw { get; set; }
        public string DiffMode { get; set; }
        public string UseWave { get; set; }
    }
}