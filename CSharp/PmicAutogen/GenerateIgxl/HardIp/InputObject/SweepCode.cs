using System;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class SweepCode
    {
        public SweepCode()
        {
            SendBitName = "";
            Width = 0;
            Start = 0;
            Step = 0;
            End = 0;
            Copy = 1;
            IsGrayCode = false;
            Order = "";
        }

        public string SendBitName { set; get; }
        public int Width { set; get; }
        public int Start { set; get; }
        public int Step { set; get; }
        public int End { set; get; }
        public int Copy { set; get; }
        public bool IsGrayCode { set; get; }
        public string Order { set; get; }
        public string Misc { set; get; }
    }
}