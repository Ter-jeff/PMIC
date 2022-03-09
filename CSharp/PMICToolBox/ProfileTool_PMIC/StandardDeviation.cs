using System;
using System.Collections.Generic;
using System.Linq;

namespace ProfileTool_PMIC
{
    public static class StandardDeviation
    {
        public static double GetStandardDeviation(this IEnumerable<double> values)
        {
            double standardDeviation = 0;
            var enumerable = values as double[] ?? values.ToArray();
            var count = enumerable.Count();
            if (count > 1)
            {
                var avg = enumerable.Average();
                var sum = enumerable.Sum(d => (d - avg) * (d - avg));
                standardDeviation = Math.Sqrt((sum / count));
            }
            return standardDeviation;
        }
    }
}