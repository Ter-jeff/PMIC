using System;
using System.Collections.Generic;
using System.Linq;

namespace AutomationCommon.Utility
{
    public class PaEngineFrc
    {
        public double TargetFreq { set; get; }
        public int M2 { set; get; }
        public double PatgenFreq { set; get; }
        public int X2 { set; get; }
        public int D2 { set; get; }
        public double ClkD8Freq { set; get; }
        public int M { set; get; }
        public double PdfFreq { set; get; }
        public int D1 { set; get; }
        public double PllInputFreq { set; get; }

        public PaEngineFrc()
        {
            TargetFreq = 0;
            M2 = 0;
            PatgenFreq = 0;
            X2 = 0;
            D2 = 0;
            ClkD8Freq = 0;
            M = 0;
            PdfFreq = 0;
            D1 = 0;
            PllInputFreq = 0;
        }
    }

    public class PaEngine
    {
        public int TargetFreq { set; get; }
        public int D2 { set; get; }
        public int ClkD8Freq { set; get; }
        public int M { set; get; }
        public int PdfFreq { set; get; }
        public int D1 { set; get; }
        public int PllInputFreq { set; get; }

        public PaEngine()
        {
            TargetFreq = 0;
            D2 = 0;
            ClkD8Freq = 0;
            M = 0;
            PdfFreq = 0;
            D1 = 0;
            PllInputFreq = 0;
        }
    }



}