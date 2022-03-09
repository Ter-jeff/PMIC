using System;
using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Utility
{
    public class FRCCalc
    {
        public List<double> CalculateFRCFreq(List<int> dataList)
        {
            var SbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();
            var TotalSetOfFreq = new List<List<double>>();
            List<String> TotalFreqRef = new List<string>();

            TotalFreqRef.Clear();
            foreach (var data in dataList)
            {

                //represent to items of nWire

                SbcFreqCalculatorFrc calculatorFRC = new SbcFreqCalculatorFrc();


                calculatorFRC.TargetFreq.Add(data);

                calculatorFRC.SolveSbcFreqFrc(out SbcFreqCalculatorFrcList);
                var CurrFreqSet = SbcFreqCalculatorFrcList.Select(p => p.Value.EngineList[0].PllInputFreq).ToList();
                TotalSetOfFreq.Add(CurrFreqSet);
            }
            //Calculate total set of frequency to search minimum frequency of total use
            var TargetSets = (from list in TotalSetOfFreq
                              from option in list
                              where TotalSetOfFreq.All(l => l.Any(o => o == option))
                              orderby option
                              select option).ToList().Distinct();
            //
            var SBCs = new List<double>();
            foreach (var pllFreq in TargetSets)
            {
                var target = SbcFreqCalculatorFrcList.FirstOrDefault(p => p.Value.EngineList[0].PllInputFreq == pllFreq);
                SBCs.Add(target.Value.SbcFreq);
            }
            return SBCs;
        }
    }
}
