using System;
using System.Collections.Generic;
using System.Linq;

namespace AutomationCommon.Utility
{
    public class FrcCalc
    {
        public List<double> CalculateFrcFreq(List<int> dataList)
        {
            var sbcFreqCalculatorFrcList = new Dictionary<string, SbcSolutionFrc>();
            var totalSetOfFreq = new List<List<double>>();
            List<String> totalFreqRef = new List<string>();

            totalFreqRef.Clear();
            foreach (var data in dataList)
            {

                //represent to items of nWire

                SbcFreqCalculatorFrc calculatorFrc = new SbcFreqCalculatorFrc();


                calculatorFrc.TargetFreq.Add(data);

                calculatorFrc.SolveSbcFreqFrc(out sbcFreqCalculatorFrcList);
                var currFreqSet = sbcFreqCalculatorFrcList.Select(p => p.Value.EngineList[0].PllInputFreq).ToList();
                totalSetOfFreq.Add(currFreqSet);
            }
            //Calculate total set of frequency to search minimum frequency of total use
            var targetSets = (from list in totalSetOfFreq
                              from option in list
                              where totalSetOfFreq.All(l => l.Any(o => o == option))
                              orderby option
                              select option).ToList().Distinct();
            //
            var sbCs = new List<double>();
            foreach (var pllFreq in targetSets)
            {
                var target = sbcFreqCalculatorFrcList.FirstOrDefault(p => p.Value.EngineList[0].PllInputFreq == pllFreq);
                sbCs.Add(target.Value.SbcFreq);
            }
            return sbCs;
        }
    }
}
