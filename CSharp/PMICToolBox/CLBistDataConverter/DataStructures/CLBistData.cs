using Library.DataStruct;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CLBistDataConverter.DataStructures
{
    public class CLBistDie
    {
        public List<CLBistSite> CLBistSitelst = new List<CLBistSite>();     
    }
    public class CLBistSite
    {
        public static string NA = "N/A";    
        public string LotId;
        public string WaferId;
        public string DieLocationX;
        public string DieLocationY;
        public string DftGroup;
        public Dictionary<string, string> clkCfg5Dic = new Dictionary<string, string>();
        public List<CLBistDataLogRow> DatalogRows = new List<CLBistDataLogRow>();

        private string _site;
        private List<CLBistSiteOutputRow> _outputRowlist;

        public string Site
        {
            get { return this._site; }
        }

        public List<CLBistSiteOutputRow> OutputRowlist
        {
            get { return _outputRowlist; }
        }
        public CLBistSite(string site)
        {
            this._site = site;
        }

        public void EditSiteClbistOutputRows()
        {
            _outputRowlist = new List<CLBistSiteOutputRow>();
            string phase,dacNumber;
            List<CLBistDataLogRow> phaseLogRows, dacLogRows;
            CLBistSiteOutputRow outputRow;
           var groupsByPhase = DatalogRows.GroupBy(p => p.Phase);
            foreach (var phaseGroup in groupsByPhase)
            {
                phase = phaseGroup.Key;
                if (string.IsNullOrEmpty(phase))
                    continue;
                phaseLogRows = phaseGroup.ToList();
                var groupsByDac = phaseLogRows.GroupBy(p => p.DacNumber);
                foreach (var dacGroup in groupsByDac)
                {
                    dacNumber = dacGroup.Key;
                    dacLogRows = dacGroup.ToList();

                    outputRow = new CLBistSiteOutputRow();
                    outputRow.ClNumber = phase;
                    outputRow.BDac1 = dacNumber;
                    outputRow.BDac2 = dacNumber;
                    outputRow.Clk_cfg5 = clkCfg5Dic.ContainsKey(dacNumber) ? clkCfg5Dic[dacNumber] : "0";
                    GetDacOutputRowData(outputRow, dacLogRows);
                    _outputRowlist.Add(outputRow);
                }
            }
        } 

        private void GetDacOutputRowData(CLBistSiteOutputRow outputRow, List<CLBistDataLogRow> dacLogRows)
        {
            string[] testNameArr;
            outputRow.Freq = NA;
            outputRow.BlRef = NA;
            outputRow.I1 = NA;
            outputRow.I2 = NA;
            outputRow.I12 = NA;
            outputRow.I21 = NA;
            outputRow.L11 = NA;
            outputRow.L22 = NA;
            outputRow.L12 = NA;
            outputRow.L21 = NA;
            outputRow.K1 = NA;
            outputRow.K2 = NA;
            outputRow.R11 = NA;
            outputRow.R22 = NA;
            outputRow.R12 = NA;
            outputRow.R21 = NA;
            outputRow.Rdc1 = NA;
            outputRow.Rdc2 = NA;
            outputRow.RrefA = NA;
            outputRow.RrefB = NA;
            outputRow.Vddh = NA;
            foreach (CLBistDataLogRow logRow in dacLogRows)
            {
                testNameArr = logRow.TestName.Split('_');
                if (outputRow.Vddh == NA && testNameArr.Length>=8 && RegStore.RegVddh.IsMatch(testNameArr[8]))
                {
                    outputRow.Vddh = RegStore.RegVddh.Match(testNameArr[8]).Groups["vddh"].ToString();
                }
                if (outputRow.Freq == NA && RegStore.RegFreq.IsMatch(testNameArr[testNameArr.Length - 2]))
                {
                    outputRow.Freq = RegStore.RegFreq.Match(testNameArr[testNameArr.Length - 2]).Groups["freq"].ToString();
                }
                if (outputRow.BlRef == NA && RegStore.RegBIRef.IsMatch(testNameArr[testNameArr.Length - 1]))
                {
                    outputRow.BlRef = RegStore.RegBIRef.Match(testNameArr[testNameArr.Length - 1]).Groups["biref"].ToString();
                }
                if (outputRow.I1 == NA && testNameArr.Length >=3 && testNameArr[2].Equals("bDAC", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.I1 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.I2 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("bDAC", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.I2 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.I12 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("bDAC-dual", StringComparison.OrdinalIgnoreCase))
                {
                    outputRow.I12 = GetMeasuredValue(logRow.Measured);
                    outputRow.I21 = outputRow.I12;
                    continue;
                }
                if (outputRow.L11 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("L", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.L11 = GetMeasuredValue(logRow.Measured);
                }
                if (outputRow.L22 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("L", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.L22 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.L12 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("M", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.L12 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.L21 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("M", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.L21 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.K1 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("K", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.K1 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.K2 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("K", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.K2 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.R11 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rac11", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.R11 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.R22 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rac22", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.R22 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.R12 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rac12", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.R12 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.R21 == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rac21", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.R21 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.Rdc1 == NA && testNameArr.Length >= 5 && testNameArr[1].Equals("Rdc", StringComparison.OrdinalIgnoreCase) && testNameArr[4].Equals("Rdc", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.Rdc1 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.Rdc2 == NA && testNameArr.Length >= 5 && testNameArr[1].Equals("Rdc", StringComparison.OrdinalIgnoreCase) && testNameArr[4].Equals("Rdc", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.Rdc2 = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.RrefA == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rref", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("A"))
                {
                    outputRow.RrefA = GetMeasuredValue(logRow.Measured);
                    continue;
                }
                if (outputRow.RrefB == NA && testNameArr.Length >= 3 && testNameArr[2].Equals("Rref", StringComparison.OrdinalIgnoreCase) && logRow.Pin.StartsWith("B"))
                {
                    outputRow.RrefB = GetMeasuredValue(logRow.Measured);
                    continue;
                }
            }

            if (outputRow.BlRef == NA)
                outputRow.BlRef = "3";

        }

        private string GetMeasuredValue(string input)
        {
            if (string.IsNullOrEmpty(input))
                return NA;
            string[] arr = Regex.Split(input, @"[\s]+");
            if (arr.Length == 1)
                return arr[0];
            if(arr.Length >= 2)
            {
                string value = arr[0];
                string unit = arr[1][0].ToString();
                double dvalue;
                if (!double.TryParse(value, out dvalue))
                    return value;
                switch (unit)
                {
                    case "m":
                        return (dvalue * 0.001).ToString();
                    case "u":
                        return (dvalue * 0.000001).ToString();
                    case "n":
                        return (dvalue * 0.000000001).ToString();
                }
                return value;
            }
            return NA;
        }
    }

    public class CLBistSiteOutputRow
    {
        public string Clk_cfg5;
        public string ClNumber;
        public string Freq;
        public string BDac1;
        public string BDac2;
        public string BlRef;
        public string I1;
        public string I12;
        public string I2;

        public string I21;
        public string L11;
        public string L22;
        public string L12;
        public string L21;
        public string K1;
        public string K2;
        public string R11;

        public string R22;
        public string R12;
        public string R21;
        public string Rdc1;
        public string Rdc2;
        public string RrefA;
        public string RrefB;
        public string L11SubL22
        {
            get {
                double l11, l22;
                if(!string.IsNullOrEmpty(L11) && !string.IsNullOrEmpty(L22))
                {
                    if (double.TryParse(L11, out l11) && double.TryParse(L22, out l22))
                        return (l11 - l22).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string L12SubL21
        {
            get
            {
                double l12, l21;
                if (!string.IsNullOrEmpty(L12) && !string.IsNullOrEmpty(L21))
                {
                    if (double.TryParse(L12, out l12) && double.TryParse(L21, out l21))
                        return (l12 - l21).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string R11SubR22
        {
            get
            {
                double r11, r22;
                if (!string.IsNullOrEmpty(R11) && !string.IsNullOrEmpty(R22))
                {
                    if (double.TryParse(R11, out r11) && double.TryParse(R22, out r22))
                        return (r11 - r22).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string R12SubR21
        {
            get
            {
                double r12, r21;
                if (!string.IsNullOrEmpty(R12) && !string.IsNullOrEmpty(R21))
                {
                    if (double.TryParse(R12, out r12) && double.TryParse(R21, out r21))
                        return (r12 - r21).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string AveLself
        {
            get
            {
                double l11, l22;
                if (!string.IsNullOrEmpty(L11) && !string.IsNullOrEmpty(L22))
                {
                    if (double.TryParse(L11, out l11) && double.TryParse(L22, out l22))
                        return ((l11 + l22)/2).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string AveLmutual
        {
            get
            {
                double l12, l21;
                if (!string.IsNullOrEmpty(L12) && !string.IsNullOrEmpty(L21))
                {
                    if (double.TryParse(L12, out l12) && double.TryParse(L21, out l21))
                        return ((l12 + l21)/2).ToString();
                }
                return CLBistSite.NA;
            }
        }
        public string Vddh;
    }
}
