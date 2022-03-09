using System;
using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Utility
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

    public class SbcSolutionFrc
    {
        public double SbcFreq { set; get; }
        public List<PaEngineFrc> EngineList { set; get; }

        public SbcSolutionFrc()
        {
            SbcFreq = 0;
            EngineList = new List<PaEngineFrc>();
        }
    }

    public class SbcSolution
    {
        public int SbcFreq { set; get; }
        public List<PaEngine> EngineList { set; get; }

        public SbcSolution()
        {
            SbcFreq = 0;
            EngineList = new List<PaEngine>();
        }
    }

    public class SbcFreqCalculator
    {
        private int _clkD8FreqLowLimit = 125000000;
        private int _clkD8FreqHighLimit = 275000000;
        private int _pdfFreqLowLimit = 3000000;
        private int _pdfFreqHighLimit = 6000000;
        private int _pllInputFreqHighLimit = 200000000;
        private int _pllInputFreqLowLimit = 3000000;

        public string Sdf { set; get; }
        public List<int> TargetFreq;

        public SbcFreqCalculator()
        {
            TargetFreq = new List<int>();
        }

        public SbcSolution SolveSbcFreq()
        {
            SbcSolution solutions = new SbcSolution();
            int clkD8Freq, pdfFreq, pllInputFreq;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;
            int pllInputLowRange, pllInputHighRange;
            int maxTarget = TargetFreq.Max();
            List<int> tryList = TargetFreq.ToList();
            tryList.Remove(maxTarget);
            clkD8LowRange = _clkD8FreqLowLimit / maxTarget;
            clkD8HighRange = _clkD8FreqHighLimit / maxTarget;
            for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
            {
                clkD8Freq = maxTarget * i;
                if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                {
                    continue;
                }
                pdfHighRange = clkD8Freq / _pdfFreqLowLimit;
                pdfLowRange = clkD8Freq / _pdfFreqHighLimit;

                for (int j = pdfLowRange; j <= pdfHighRange; j++)
                {
                    pdfFreq = clkD8Freq / j;

                    if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                    {
                        continue;
                    }
                    pllInputLowRange = _pllInputFreqLowLimit / pdfFreq;
                    pllInputHighRange = _pllInputFreqHighLimit / pdfFreq;
                    for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                    {
                        pllInputFreq = pdfFreq * k;
                        if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                        {
                            continue;
                        }
                        List<PaEngine> engines;
                        if (IsOk(pllInputFreq, tryList, out engines))
                        {
                            SbcSolution solution = new SbcSolution();
                            solution.EngineList.AddRange(engines);
                            solution.SbcFreq = pllInputFreq * 4;
                            PaEngine engine = new PaEngine();
                            engine.ClkD8Freq = clkD8Freq;
                            engine.D2 = i;
                            engine.PdfFreq = pdfFreq;
                            engine.M = j;
                            engine.PllInputFreq = pllInputFreq;
                            engine.D1 = k;
                            engine.TargetFreq = maxTarget;
                            solution.EngineList.Add(engine);

                            return solution;
                        }
                    }
                }
            }

            SbcSolution solutionNull = new SbcSolution();
            solutionNull.SbcFreq = 62500000;
            // 32Hz to 62.5MHz.
            return solutionNull;
        }

        private bool IsOk(int tryValue, List<int> targetList, out List<PaEngine> engines)
        {
            bool suitable = true;
            engines = new List<PaEngine>();
            foreach (int target in targetList)
            {
                PaEngine engine;
                if (TrySingelValue(target, tryValue, out engine))
                {
                    engines.Add(engine);
                }
                else
                {
                    suitable = false;
                }
            }
            return suitable;
        }

        private bool TrySingelValue(int targetValue, int pllInputFreq, out PaEngine engine)
        {
            if (pllInputFreq.Equals(7680000))
            {
            }
            engine = new PaEngine();
            int clkD8Freq, pdfFreq;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;
            clkD8LowRange = _clkD8FreqLowLimit / targetValue;
            clkD8HighRange = _clkD8FreqHighLimit / targetValue;
            for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
            {
                clkD8Freq = targetValue * i;
                if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                {
                    continue;
                }
                pdfHighRange = clkD8Freq / _pdfFreqLowLimit;
                pdfLowRange = clkD8Freq / _pdfFreqHighLimit;
                if (clkD8Freq.Equals(245760000))
                {
                }
                for (int j = pdfLowRange; j <= pdfHighRange; j++)
                {

                    pdfFreq = clkD8Freq / j;
                    if (pdfFreq.Equals(3840000))
                    {
                    }
                    if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                    {
                        continue;
                    }
                    if (pllInputFreq % pdfFreq == 0)
                    {
                        engine.ClkD8Freq = clkD8Freq;
                        engine.D2 = i;
                        engine.PdfFreq = pdfFreq;
                        engine.PllInputFreq = pllInputFreq;
                        engine.M = j;
                        engine.D1 = pllInputFreq / pdfFreq;
                        engine.TargetFreq = targetValue;
                        return true;
                    }
                }
            }
            return false;
        }
    }

    public class SbcFreqCalculatorFrcList
    {
        public Dictionary<string, SbcSolutionFrc> SbcSolutionFrcList { set; get; }
    }



    public class SbcFreqCalculatorFrc
    {
        private double _patgenFreqHighLimit = 550000000;
        private double _patgenFreqLowLimit = 1907;
        private double _clkD8FreqLowLimit = 125000000;
        private double _clkD8FreqHighLimit = 275000000;
        private double _pdfFreqLowLimit = 3000000;
        private double _pdfFreqHighLimit = 6000000;
        private double _pllInputFreqHighLimit = 200000000;
        private double _pllInputFreqLowLimit = 3000000;

        public string Sdf { set; get; }
        public List<double> TargetFreq;

        public SbcFreqCalculatorFrc()
        {
            TargetFreq = new List<double>();
        }

        private bool TrySingelValue(double targetValue, double pllInputFreq, out PaEngineFrc engine)
        {

            engine = new PaEngineFrc();

            //int patgenFreq = 0;
            double patgenFreq = 0.0;
            double clkD8Freq, pdfFreq;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;

            int M2 = 0;

            for (int HSpeedMode = 1; HSpeedMode <= 2; HSpeedMode++)
            {
                if (HSpeedMode == 1)
                {
                    patgenFreq = targetValue;
                }
                else if (HSpeedMode == 2)
                {
                    patgenFreq = targetValue / 2;
                }

                if (patgenFreq >= _patgenFreqHighLimit)
                {
                    continue;
                }

                M2 = HSpeedMode;

                for (int X2 = 1; X2 <= 2; X2++)
                {
                    if (X2 == 1)
                    {
                        clkD8LowRange = (int)(_clkD8FreqLowLimit / patgenFreq);
                        clkD8HighRange = (int)(_clkD8FreqHighLimit / patgenFreq);
                        // D2 = clkD8LowRange ~ clkD8HighRange
                        for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                        {
                            clkD8Freq = patgenFreq * i;
                            if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit || i > 65535)
                            {
                                continue;
                            }
                            pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                            pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                            for (int j = pdfLowRange; j <= pdfHighRange; j++)
                            {
                                pdfFreq = clkD8Freq / j;
                                if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                // ToDo why clkD8Freq % j != 0 ??
                                //if (pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                {
                                    continue;
                                }
                                if (pllInputFreq % pdfFreq == 0)
                                {
                                    engine.ClkD8Freq = clkD8Freq;
                                    engine.M2 = M2;
                                    engine.PatgenFreq = patgenFreq;
                                    engine.D2 = i;
                                    engine.X2 = X2;
                                    engine.PdfFreq = pdfFreq;
                                    engine.PllInputFreq = pllInputFreq;
                                    engine.M = j;
                                    engine.D1 = (int)(pllInputFreq / pdfFreq);
                                    engine.TargetFreq = targetValue;
                                    return true;
                                }
                            }
                        }

                    }
                    else if (X2 == 2) // D2 only can put 1
                    {
                        clkD8Freq = patgenFreq / 2;
                        // D2 = 1;
                        if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                        {
                            continue;
                        }
                        pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                        pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                        for (int j = pdfLowRange; j <= pdfHighRange; j++)
                        {
                            pdfFreq = clkD8Freq / j;
                            if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            //if (pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            {
                                continue;
                            }
                            if (pllInputFreq % pdfFreq == 0)
                            {
                                engine.ClkD8Freq = clkD8Freq;
                                engine.M2 = M2;
                                engine.PatgenFreq = patgenFreq;
                                engine.D2 = 1;
                                engine.X2 = X2;
                                engine.PdfFreq = pdfFreq;
                                engine.PllInputFreq = pllInputFreq;
                                engine.M = j;
                                engine.D1 = (int)(pllInputFreq / pdfFreq);
                                engine.TargetFreq = targetValue;
                                return true;
                            }
                        }

                    }
                }
            }
            return false;
        }

        private bool IsOkFrc(double tryValue, List<double> targetList, out List<PaEngineFrc> engines)
        {
            bool suitable = true;
            engines = new List<PaEngineFrc>();
            foreach (double target in targetList)
            {
                PaEngineFrc engine;
                if (TrySingelValue(target, tryValue, out engine))
                {
                    engines.Add(engine);
                }
                else
                {
                    suitable = false;
                }
            }
            return suitable;
        }

        public SbcSolutionFrc SolveSbcFreqFrcCheck(string logpath, out Dictionary<string, SbcSolutionFrc> SbcFreqCalculatorFrcList, double ref_pllInputFreq)
        {
            SbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();

            SbcSolutionFrc solutions = new SbcSolutionFrc();

            int SolutionGroupNo = 0;

            double patgenFreq = 0.0;
            double clkD8Freq, pdfFreq, pllInputFreq;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;
            int pllInputLowRange, pllInputHighRange;

            //int maxTarget = TargetFreq.Max();

            double firstTarget = TargetFreq[0];

            List<double> tryList = TargetFreq.ToList();

            tryList.Remove(TargetFreq[0]);


            int M2 = 0;

            for (int HSpeedMode = 1; HSpeedMode <= 2; HSpeedMode++)
            {
                if (HSpeedMode == 1)
                {
                    patgenFreq = firstTarget;
                }
                else if (HSpeedMode == 2)
                {
                    patgenFreq = firstTarget / 2;
                }

                if (patgenFreq >= _patgenFreqHighLimit)
                {
                    continue;
                }

                M2 = HSpeedMode;


                for (int X2 = 1; X2 <= 2; X2++)
                {
                    if (X2 == 1)
                    {
                        clkD8LowRange = (int)(_clkD8FreqLowLimit / patgenFreq);
                        clkD8HighRange = (int)(_clkD8FreqHighLimit / patgenFreq);
                        // D2 = clkD8LowRange ~ clkD8HighRange
                        for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                        {
                            clkD8Freq = patgenFreq * i;
                            if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                            {
                                continue;
                            }
                            pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                            pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                            for (int j = pdfLowRange; j <= pdfHighRange; j++)
                            {
                                pdfFreq = clkD8Freq / j;
                                if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                {
                                    continue;
                                }
                                pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                                pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                                for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                                {
                                    pllInputFreq = pdfFreq * k;
                                    if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                    {
                                        continue;
                                    }

                                    if (pllInputFreq == ref_pllInputFreq)
                                    {
                                        //List<PaEngineFrc> engines;
                                        //if (IsOkFrc(pllInputFreq, tryList, out engines))
                                        //{
                                        SbcSolutionFrc solution = new SbcSolutionFrc();
                                        //solution.EngineList.AddRange(engines);
                                        solution.SbcFreq = pllInputFreq * 4;
                                        PaEngineFrc engine = new PaEngineFrc();
                                        engine.ClkD8Freq = clkD8Freq;
                                        engine.M2 = M2;
                                        engine.PatgenFreq = patgenFreq;
                                        engine.D2 = i;
                                        engine.X2 = X2;
                                        engine.PdfFreq = pdfFreq;
                                        engine.M = j;
                                        engine.PllInputFreq = pllInputFreq;
                                        engine.D1 = k;
                                        engine.TargetFreq = firstTarget;
                                        solution.EngineList.Add(engine);
                                        SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                        SolutionGroupNo++;
                                    }



                                    //}
                                }
                            }
                        }

                    }
                    else if (X2 == 2) // D2 only can put 1
                    {
                        clkD8Freq = patgenFreq / 2;
                        // D2 = 1;
                        if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                        {
                            continue;
                        }
                        pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                        pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                        for (int j = pdfLowRange; j <= pdfHighRange; j++)
                        {
                            pdfFreq = clkD8Freq / j;
                            if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            {
                                continue;
                            }
                            pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                            pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                            for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                            {
                                pllInputFreq = pdfFreq * k;
                                if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                {
                                    continue;
                                }
                                //List<PaEngineFrc> engines;
                                //if (IsOkFrc(pllInputFreq, tryList, out engines))
                                //{

                                if (pllInputFreq == ref_pllInputFreq)
                                {
                                    SbcSolutionFrc solution = new SbcSolutionFrc();
                                    //solution.EngineList.AddRange(engines);
                                    solution.SbcFreq = pllInputFreq * 4;
                                    PaEngineFrc engine = new PaEngineFrc();
                                    engine.ClkD8Freq = clkD8Freq;
                                    engine.M2 = M2;
                                    engine.PatgenFreq = patgenFreq;
                                    engine.D2 = 1;
                                    engine.X2 = X2;
                                    engine.PdfFreq = pdfFreq;
                                    engine.M = j;
                                    engine.PllInputFreq = pllInputFreq;
                                    engine.D1 = k;
                                    engine.TargetFreq = firstTarget;
                                    solution.EngineList.Add(engine);

                                    SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                    SolutionGroupNo++;
                                }
                                //}
                            }
                        }
                    }
                }
            }

            if (SbcFreqCalculatorFrcList.Count > 0)
            {
                return SbcFreqCalculatorFrcList["0"];
            }
            else
            {
                SbcSolutionFrc solutionNull = new SbcSolutionFrc();
                solutionNull.SbcFreq = 62500000;
                // 32Hz to 62.5MHz.
                return solutionNull;
            }
        }

        public SbcSolutionFrc SolveSbcFreqFrcCheckInvert(string logpath, out Dictionary<string, SbcSolutionFrc> SbcFreqCalculatorFrcList, double ref_pllInputFreq)
        {
            SbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();

            SbcSolutionFrc solutions = new SbcSolutionFrc();

            int SolutionGroupNo = 0;

            double patgenFreq;
            double TargetSolFreq;
            double clkD8Freq, pdfFreq;
            int patgenHighRange;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;

            //int maxTarget = TargetFreq.Max();

            double firstTarget = TargetFreq[0];

            List<double> tryList = TargetFreq.ToList();

            tryList.Remove(TargetFreq[0]);

            pdfHighRange = (int)(ref_pllInputFreq / _pdfFreqLowLimit);
            pdfLowRange = (int)(ref_pllInputFreq / _pdfFreqHighLimit);


            List<string> totalSollist = new List<string>();
            totalSollist.Clear();

            for (int j = pdfLowRange; j <= pdfHighRange; j++)
            {
                pdfFreq = ref_pllInputFreq / (double)j;
                if (ref_pllInputFreq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                {
                    continue;
                }

                clkD8HighRange = (int)(_clkD8FreqHighLimit / pdfFreq);
                clkD8LowRange = (int)(_clkD8FreqLowLimit / pdfFreq);

                for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                {
                    clkD8Freq = pdfFreq * i;
                    if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                    {
                        continue;
                    }

                    for (int X2 = 1; X2 <= 2; X2++)
                    {
                        for (int HSpeedMode = 2; HSpeedMode >= 1; HSpeedMode--)
                        {
                            if (HSpeedMode == 1)   // normal speed
                            {
                                if (X2 == 1)
                                {
                                    //patgenLowRange = (int)(clkD8Freq / _patgenFreqHighLimit);
                                    patgenHighRange = (int)(clkD8Freq / Math.Max(firstTarget, _patgenFreqLowLimit));

                                    if (patgenHighRange > 65535)
                                        patgenHighRange = 65535;


                                    for (int k = 1; k <= patgenHighRange; k++)
                                    {
                                        patgenFreq = clkD8Freq / k;

                                        if (clkD8Freq % k != 0 || patgenFreq >= _patgenFreqHighLimit || patgenFreq <= _patgenFreqLowLimit)
                                        {
                                            continue;
                                        }

                                        TargetSolFreq = patgenFreq;

                                        if (firstTarget <= TargetSolFreq * 0.80 || firstTarget >= TargetSolFreq * 1.2)
                                        {
                                            continue;
                                        }

                                        SbcSolutionFrc solution = new SbcSolutionFrc();
                                        solution.SbcFreq = ref_pllInputFreq * 4;
                                        PaEngineFrc engine = new PaEngineFrc();
                                        engine.ClkD8Freq = clkD8Freq;
                                        engine.M2 = HSpeedMode;
                                        engine.PatgenFreq = patgenFreq;
                                        engine.D2 = k;
                                        engine.X2 = X2;
                                        engine.PdfFreq = pdfFreq;
                                        engine.M = i;
                                        engine.PllInputFreq = ref_pllInputFreq;
                                        engine.D1 = j;
                                        engine.TargetFreq = TargetSolFreq;
                                        solution.EngineList.Add(engine);


                                        if (!totalSollist.Contains(TargetSolFreq.ToString()))
                                            totalSollist.Add(TargetSolFreq.ToString());
                                        else
                                        {
                                            continue;
                                        }

                                        SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                        SolutionGroupNo++;
                                    }
                                }
                                else if (X2 == 2)
                                {
                                    patgenFreq = clkD8Freq * 2;
                                    if (patgenFreq >= _patgenFreqHighLimit || patgenFreq <= _patgenFreqLowLimit)
                                    {
                                        continue;
                                    }

                                    TargetSolFreq = patgenFreq;

                                    if (firstTarget <= TargetSolFreq * 0.8 || firstTarget >= TargetSolFreq * 1.2)
                                    {
                                        continue;
                                    }

                                    SbcSolutionFrc solution = new SbcSolutionFrc();
                                    solution.SbcFreq = ref_pllInputFreq * 4;
                                    PaEngineFrc engine = new PaEngineFrc();
                                    engine.ClkD8Freq = clkD8Freq;
                                    engine.M2 = HSpeedMode;
                                    engine.PatgenFreq = patgenFreq;
                                    engine.D2 = 1;
                                    engine.X2 = X2;
                                    engine.PdfFreq = pdfFreq;
                                    engine.M = i;
                                    engine.PllInputFreq = ref_pllInputFreq;
                                    engine.D1 = j;
                                    engine.TargetFreq = TargetSolFreq;
                                    solution.EngineList.Add(engine);

                                    if (!totalSollist.Contains(TargetSolFreq.ToString()))
                                        totalSollist.Add(TargetSolFreq.ToString());
                                    else
                                    {
                                        continue;
                                    }


                                    SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                    SolutionGroupNo++;
                                }

                            }
                            else if (HSpeedMode == 2)  // HighSpeed Mode
                            {
                                if (X2 == 1)
                                {
                                    //patgenLowRange = (int)(clkD8Freq / _patgenFreqHighLimit);
                                    patgenHighRange = (int)(clkD8Freq / Math.Max(firstTarget, _patgenFreqLowLimit));

                                    if (patgenHighRange > 65535)
                                        patgenHighRange = 65535;

                                    for (int k = 1; k <= patgenHighRange; k++)
                                    {
                                        patgenFreq = clkD8Freq / k;

                                        if (clkD8Freq % k != 0 || patgenFreq >= _patgenFreqHighLimit || patgenFreq <= _clkD8FreqLowLimit)
                                        {
                                            continue;
                                        }

                                        TargetSolFreq = patgenFreq * 2;

                                        if (firstTarget <= TargetSolFreq * 0.8 || firstTarget >= TargetSolFreq * 1.2)
                                        {
                                            continue;
                                        }

                                        SbcSolutionFrc solution = new SbcSolutionFrc();
                                        solution.SbcFreq = ref_pllInputFreq * 4;
                                        PaEngineFrc engine = new PaEngineFrc();
                                        engine.ClkD8Freq = clkD8Freq;
                                        engine.M2 = HSpeedMode;
                                        engine.PatgenFreq = patgenFreq;
                                        engine.D2 = k;
                                        engine.X2 = X2;
                                        engine.PdfFreq = pdfFreq;
                                        engine.M = i;
                                        engine.PllInputFreq = ref_pllInputFreq;
                                        engine.D1 = j;
                                        engine.TargetFreq = TargetSolFreq;
                                        solution.EngineList.Add(engine);


                                        if (!totalSollist.Contains(TargetSolFreq.ToString()))
                                            totalSollist.Add(TargetSolFreq.ToString());
                                        else
                                        {
                                            continue;
                                        }

                                        SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                        SolutionGroupNo++;
                                    }
                                }
                                else
                                {
                                    patgenFreq = clkD8Freq * 2;
                                    if (patgenFreq >= _patgenFreqHighLimit || patgenFreq <= _clkD8FreqLowLimit)
                                    {
                                        continue;
                                    }

                                    TargetSolFreq = patgenFreq * 2;

                                    if (firstTarget <= TargetSolFreq * 0.8 || firstTarget >= TargetSolFreq * 1.2)
                                    {
                                        continue;
                                    }

                                    SbcSolutionFrc solution = new SbcSolutionFrc();
                                    solution.SbcFreq = ref_pllInputFreq * 4;
                                    PaEngineFrc engine = new PaEngineFrc();
                                    engine.ClkD8Freq = clkD8Freq;
                                    engine.M2 = HSpeedMode;
                                    engine.PatgenFreq = patgenFreq;
                                    engine.D2 = 1;
                                    engine.X2 = X2;
                                    engine.PdfFreq = pdfFreq;
                                    engine.M = i;
                                    engine.PllInputFreq = ref_pllInputFreq;
                                    engine.D1 = j;
                                    engine.TargetFreq = TargetSolFreq;
                                    solution.EngineList.Add(engine);

                                    if (!totalSollist.Contains(TargetSolFreq.ToString()))
                                        totalSollist.Add(TargetSolFreq.ToString());
                                    else
                                    {
                                        continue;
                                    }

                                    SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                    SolutionGroupNo++;
                                }
                            }
                        }
                    }
                }
            }












            /*

                        int M2 = 0;

                        for (int HSpeedMode = 1; HSpeedMode <= 2; HSpeedMode++)
                        {
                            if (HSpeedMode == 1)
                            {
                                patgenFreq = firstTarget;
                            }
                            else if (HSpeedMode == 2)
                            {
                                patgenFreq = firstTarget / 2;
                            }

                            if (patgenFreq >= _patgenFreqHighLimit)
                            {
                                continue;
                            }

                            M2 = HSpeedMode;


                            for (int X2 = 1; X2 <= 2; X2++)
                            {
                                if (X2 == 1)
                                {
                                    clkD8LowRange = (int)(_clkD8FreqLowLimit / patgenFreq);
                                    clkD8HighRange = (int)(_clkD8FreqHighLimit / patgenFreq);
                                    // D2 = clkD8LowRange ~ clkD8HighRange
                                    for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                                    {
                                        clkD8Freq = patgenFreq * i;
                                        if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                                        {
                                            continue;
                                        }
                                        pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                                        pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                                        for (int j = pdfLowRange; j <= pdfHighRange; j++)
                                        {
                                            pdfFreq = clkD8Freq / j;
                                            if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                            {
                                                continue;
                                            }
                                            pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                                            pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                                            for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                                            {
                                                pllInputFreq = pdfFreq * k;
                                                if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                                {
                                                    continue;
                                                }

                                                if (pllInputFreq == ref_pllInputFreq)
                                                {
                                                    //List<PaEngineFrc> engines;
                                                    //if (IsOkFrc(pllInputFreq, tryList, out engines))
                                                    //{
                                                    SbcSolutionFrc solution = new SbcSolutionFrc();
                                                    //solution.EngineList.AddRange(engines);
                                                    solution.SbcFreq = pllInputFreq * 4;
                                                    PaEngineFrc engine = new PaEngineFrc();
                                                    engine.ClkD8Freq = clkD8Freq;
                                                    engine.M2 = M2;
                                                    engine.PatgenFreq = patgenFreq;
                                                    engine.D2 = i;
                                                    engine.X2 = X2;
                                                    engine.PdfFreq = pdfFreq;
                                                    engine.M = j;
                                                    engine.PllInputFreq = pllInputFreq;
                                                    engine.D1 = k;
                                                    engine.TargetFreq = firstTarget;
                                                    solution.EngineList.Add(engine);
                                                    SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                                    SolutionGroupNo++;
                                                }



                                                //}
                                            }
                                        }
                                    }

                                }
                                else if (X2 == 2) // D2 only can put 1
                                {
                                    clkD8Freq = patgenFreq / 2;
                                    // D2 = 1;
                                    if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                                    {
                                        continue;
                                    }
                                    pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                                    pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                                    for (int j = pdfLowRange; j <= pdfHighRange; j++)
                                    {
                                        pdfFreq = clkD8Freq / j;
                                        if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                        {
                                            continue;
                                        }
                                        pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                                        pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                                        for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                                        {
                                            pllInputFreq = pdfFreq * k;
                                            if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                            {
                                                continue;
                                            }
                                            //List<PaEngineFrc> engines;
                                            //if (IsOkFrc(pllInputFreq, tryList, out engines))
                                            //{

                                            if (pllInputFreq == ref_pllInputFreq)
                                            {
                                                SbcSolutionFrc solution = new SbcSolutionFrc();
                                                //solution.EngineList.AddRange(engines);
                                                solution.SbcFreq = pllInputFreq * 4;
                                                PaEngineFrc engine = new PaEngineFrc();
                                                engine.ClkD8Freq = clkD8Freq;
                                                engine.M2 = M2;
                                                engine.PatgenFreq = patgenFreq;
                                                engine.D2 = 1;
                                                engine.X2 = X2;
                                                engine.PdfFreq = pdfFreq;
                                                engine.M = j;
                                                engine.PllInputFreq = pllInputFreq;
                                                engine.D1 = k;
                                                engine.TargetFreq = firstTarget;
                                                solution.EngineList.Add(engine);

                                                SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                                SolutionGroupNo++;
                                            }
                                            //}
                                        }
                                    }
                                }
                            }
                        }
            */
            if (SbcFreqCalculatorFrcList.Count > 0)
            {
                return SbcFreqCalculatorFrcList["0"];
            }
            else
            {
                SbcSolutionFrc solutionNull = new SbcSolutionFrc();
                solutionNull.SbcFreq = 62500000;
                // 32Hz to 62.5MHz.
                return solutionNull;
            }
        }


        public SbcSolutionFrc SolveSbcFreqFrc(out Dictionary<string, SbcSolutionFrc> SbcFreqCalculatorFrcList)
        {

            SbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();

            SbcSolutionFrc solutions = new SbcSolutionFrc();

            int SolutionGroupNo = 0;

            double patgenFreq = 0.0;
            double clkD8Freq, pdfFreq, pllInputFreq;
            int clkD8LowRange, clkD8HighRange;
            int pdfLowRange, pdfHighRange;
            int pllInputLowRange, pllInputHighRange;

            //double maxTarget = TargetFreq.Max();

            double firstTarget = TargetFreq[0];

            List<double> tryList = TargetFreq.ToList();

            tryList.Remove(TargetFreq[0]);
            //tryList.Remove(maxTarget);


            List<double> existData = new List<double>();
            existData.Clear();

            //DateTime Date = DateTime.Now;
            //string TodyTime = Date.ToString("yyyy-MM-dd-HH:mm:ss");
            //File.AppendAllText(logpath, "1st Target:" + firstTarget.ToString());
            int M2 = 0;

            for (int HSpeedMode = 1; HSpeedMode <= 2; HSpeedMode++)
            {

                if (firstTarget <= 250000000)
                {
                    if (HSpeedMode == 2) continue;
                }

                if (firstTarget >= 550000000)
                {
                    if (HSpeedMode == 1) continue;
                }



                if (HSpeedMode == 1)
                {
                    patgenFreq = firstTarget;
                }
                else if (HSpeedMode == 2)
                {
                    patgenFreq = firstTarget / 2;
                }

                if (patgenFreq >= _patgenFreqHighLimit)
                {
                    continue;
                }

                M2 = HSpeedMode;


                for (int X2 = 1; X2 <= 2; X2++)
                {
                    if (X2 == 1)
                    {
                        clkD8LowRange = (int)(_clkD8FreqLowLimit / patgenFreq);
                        clkD8HighRange = (int)(_clkD8FreqHighLimit / patgenFreq);
                        // D2 = clkD8LowRange ~ clkD8HighRange
                        for (int i = clkD8LowRange; i <= clkD8HighRange; i++)
                        {


                            clkD8Freq = patgenFreq * i;
                            if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                            {
                                continue;
                            }
                            pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                            pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                            for (int j = pdfLowRange; j <= pdfHighRange; j++)
                            {

                                pdfFreq = clkD8Freq / j;
                                if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                //if (pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                                {
                                    continue;
                                }
                                pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                                pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                                for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                                {

                                    pllInputFreq = pdfFreq * k;
                                    if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                    {
                                        continue;
                                    }
                                    List<PaEngineFrc> engines;
                                    if (IsOkFrc(pllInputFreq, tryList, out engines))
                                    {
                                        SbcSolutionFrc solution = new SbcSolutionFrc();
                                        solution.EngineList.AddRange(engines);
                                        solution.SbcFreq = pllInputFreq * 4;
                                        PaEngineFrc engine = new PaEngineFrc();
                                        engine.ClkD8Freq = clkD8Freq;
                                        engine.M2 = M2;
                                        engine.PatgenFreq = patgenFreq;
                                        engine.D2 = i;
                                        engine.X2 = X2;
                                        engine.PdfFreq = pdfFreq;
                                        engine.M = j;
                                        engine.PllInputFreq = pllInputFreq;
                                        engine.D1 = k;
                                        engine.TargetFreq = firstTarget;
                                        solution.EngineList.Add(engine);


                                        //File.AppendAllText(logpath, "\n " + SolutionGroupNo.ToString() + " 共同解");
                                        //string header = "Target Frequency\tM2\tPatgen Frequency\tD2\t*2\tClkD8Freq (>125MHz, <275MHz)\tM\tPDFFreq\tD1\tPllInputFreq\tRefClkFreq (4*PllInputFreq)";
                                        //File.AppendAllText(logpath, "\n" + header);

                                        //foreach (var item in solution.EngineList)
                                        //{
                                        //    string content = "\n";
                                        //    content += item.TargetFreq + "\t" + item.M2 + "\t" + item.PatgenFreq + "\t" + item.D2 + "\t" + item.X2 + "\t";
                                        //    content += item.ClkD8Freq + "\t" + item.M + "\t" + item.PdfFreq + "\t" + item.D1 + "\t" + item.PllInputFreq + "\t" + item.PllInputFreq * 4;
                                        //    File.AppendAllText(logpath, content);
                                        //}

                                        if (existData.Contains(pllInputFreq)) continue;
                                        else
                                        {
                                            existData.Add(pllInputFreq);
                                        }




                                        SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                        SolutionGroupNo++;

                                        //return solution;
                                    }
                                }
                            }
                        }

                    }
                    else if (X2 == 2) // D2 only can put 1
                    {
                        clkD8Freq = patgenFreq / 2;
                        // D2 = 1;
                        if (clkD8Freq <= _clkD8FreqLowLimit || clkD8Freq >= _clkD8FreqHighLimit)
                        {
                            continue;
                        }
                        pdfHighRange = (int)(clkD8Freq / _pdfFreqLowLimit);
                        pdfLowRange = (int)(clkD8Freq / _pdfFreqHighLimit);
                        for (int j = pdfLowRange; j <= pdfHighRange; j++)
                        {
                            pdfFreq = clkD8Freq / j;
                            if (clkD8Freq % j != 0 || pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            //if (pdfFreq >= _pdfFreqHighLimit || pdfFreq <= _pdfFreqLowLimit)
                            {
                                continue;
                            }
                            pllInputLowRange = (int)(_pllInputFreqLowLimit / pdfFreq);
                            pllInputHighRange = (int)(_pllInputFreqHighLimit / pdfFreq);
                            for (int k = pllInputLowRange; k <= pllInputHighRange; k++)
                            {
                                pllInputFreq = pdfFreq * k;
                                if (pllInputFreq >= _pllInputFreqHighLimit || pllInputFreq <= _pllInputFreqLowLimit)
                                {
                                    continue;
                                }
                                List<PaEngineFrc> engines;
                                if (IsOkFrc(pllInputFreq, tryList, out engines))
                                {
                                    SbcSolutionFrc solution = new SbcSolutionFrc();
                                    solution.EngineList.AddRange(engines);
                                    solution.SbcFreq = pllInputFreq * 4;
                                    PaEngineFrc engine = new PaEngineFrc();
                                    engine.ClkD8Freq = clkD8Freq;
                                    engine.M2 = M2;
                                    engine.PatgenFreq = patgenFreq;
                                    engine.D2 = 1;
                                    engine.X2 = X2;
                                    engine.PdfFreq = pdfFreq;
                                    engine.M = j;
                                    engine.PllInputFreq = pllInputFreq;
                                    engine.D1 = k;
                                    engine.TargetFreq = firstTarget;
                                    solution.EngineList.Add(engine);

                                    //File.AppendAllText(logpath, "\n " + SolutionGroupNo.ToString() + " 共同解");
                                    //string header = "Target Frequency\tM2\tPatgen Frequency\tD2\t*2\tClkD8Freq (>125MHz, <275MHz)\tM\tPDFFreq\tD1\tPllInputFreq\tRefClkFreq (4*PllInputFreq)";
                                    //File.AppendAllText(logpath, "\n" + header);

                                    //foreach (var item in solution.EngineList)
                                    //{
                                    //    string content = "\n";
                                    //    content += item.TargetFreq + "\t" + item.M2 + "\t" + item.PatgenFreq + "\t" + item.D2 + "\t" + item.X2 + "\t";
                                    //    content += item.ClkD8Freq + "\t" + item.M + "\t" + item.PdfFreq + "\t" + item.D1 + "\t" + item.PllInputFreq + "\t" + item.PllInputFreq * 4;
                                    //    File.AppendAllText(logpath, content);
                                    //}

                                    if (existData.Contains(pllInputFreq)) continue;
                                    else
                                    {
                                        existData.Add(pllInputFreq);
                                    }


                                    SbcFreqCalculatorFrcList.Add(SolutionGroupNo.ToString(), solution);
                                    SolutionGroupNo++;

                                }
                            }
                        }
                    }
                }
            }

            if (SbcFreqCalculatorFrcList.Count > 0)
            {
                return SbcFreqCalculatorFrcList["0"];
            }
            else
            {
                SbcSolutionFrc solutionNull = new SbcSolutionFrc();
                solutionNull.SbcFreq = 62500000;
                // 32Hz to 62.5MHz.
                return solutionNull;
            }



        }
    }
}