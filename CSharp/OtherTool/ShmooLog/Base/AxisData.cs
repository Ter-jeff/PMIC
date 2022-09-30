using System;
using System.Collections.Generic;
using System.Linq;

namespace ShmooLog.Base
{
    public class AxisData
    {
        public string AxisName;
        public Dictionary<string, List<double>> PinValueSet = new Dictionary<string, List<double>>();
        public string Unit = "";


        public AxisData(string axisName)
        {
            AxisName = axisName;
        }

        public void PointUnitConvert()
        {
            var lBased = 0.0;
            var lUnit = "";
            if (PinValueSet.Any())
            {
                var checkValue = PinValueSet.First().Value.FirstOrDefault();

                if (checkValue > 1)
                    return;

                if (0.001 > checkValue && checkValue > 0.000001) // u
                {
                    lUnit = "u";
                    lBased = 0.000001;
                }
                else if (0.000001 > checkValue && checkValue > 0.000000001) // n
                {
                    lUnit = "n";
                    lBased = 0.000000001;
                }
                else if (0.000000001 > checkValue && checkValue > 0.000000000001) // p 
                {
                    lUnit = "p";
                    lBased = 0.000000000001;
                }
                else if (0.000000000001 > checkValue && checkValue > 0.000000000000001) // f 
                {
                    lUnit = "f";
                    lBased = 0.000000000000001;
                }
                else
                {
                    return;
                }

                var reNew = new Dictionary<string, List<double>>();
                foreach (var pinItem in PinValueSet)
                {
                    if (!reNew.ContainsKey(pinItem.Key))
                        reNew.Add(pinItem.Key, new List<double>());
                    foreach (var value in pinItem.Value)
                        reNew[pinItem.Key].Add(Math.Round(value / lBased, 3, MidpointRounding.AwayFromZero));
                    // Math.Round(startPoint + currStepSize * i, 3,MidpointRounding.AwayFromZero)
                }

                PinValueSet = reNew;
                Unit = lUnit;
            }
        }
    }
}