using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.Relay.Input
{
    public class ADG1414Group
    {
        public string DesignNameOrg;
        public string DesignName;
        public string[] SNames = new string[8];
        public string[] SNamesOrg = new string[8];
        public string[] DNames = new string[8];
        public string[] DNamesOrg = new string[8];
        public string SDO;
    }

    public class Adg1414Reader
    {
        private static Regex _regADG1414 = new Regex("^(ADG1414)", RegexOptions.IgnoreCase);
        private static Regex _regInPin = new Regex(@"[S](?<num>\d+)$", RegexOptions.IgnoreCase);
        private static Regex _regInType = new Regex(@"[D](?<num>\d+)$", RegexOptions.IgnoreCase);
        private static Regex _regSNamePat1 = new Regex(@"^[S]\d[_]", RegexOptions.IgnoreCase);
        private static Regex _regSNamePat2 = new Regex(@"^\d[\.]", RegexOptions.IgnoreCase);

        public List<ADG1414Group> ReadData(List<ComPinRow> comPinRows)
        {
            List<ADG1414Group> ADG1414Data = new List<ADG1414Group>();
            foreach (ComPinRow comPinRow in comPinRows)
            {
                if (!_regADG1414.IsMatch(comPinRow.CompDeviceType))
                {
                    continue;
                }

                ADG1414Group adg1414Group = ADG1414Data.FirstOrDefault(adg1414 => adg1414.DesignNameOrg == comPinRow.Refdes);
                if (adg1414Group == null)
                {
                    adg1414Group = new ADG1414Group() { DesignName = ConvertDesignName(comPinRow.Refdes), DesignNameOrg = comPinRow.Refdes };
                    ADG1414Data.Add(adg1414Group);
                }

                if (_regInPin.IsMatch(comPinRow.PinName))
                {
                    int inPinNum = -1;
                    Match match = _regInPin.Match(comPinRow.PinName);
                    inPinNum = int.Parse(match.Groups["num"].Value);

                    if (inPinNum <= 8)
                    {
                        adg1414Group.SNamesOrg[inPinNum - 1] = comPinRow.NetName;
                        adg1414Group.SNames[inPinNum - 1] = ConvertSName(comPinRow.NetName);
                        if (adg1414Group.SNames[inPinNum - 1] != "") adg1414Group.SNames[inPinNum - 1] += "_S" + inPinNum.ToString();
                    }
                }
                else if (_regInType.IsMatch(comPinRow.PinName))
                {
                    int inPinNum = -1;
                    Match match = _regInType.Match(comPinRow.PinName);
                    inPinNum = int.Parse(match.Groups["num"].Value);

                    if (inPinNum <= 8)
                    {
                        adg1414Group.DNamesOrg[inPinNum - 1] = comPinRow.NetName;
                        adg1414Group.DNames[inPinNum - 1] = ConvertSName(comPinRow.NetName);
                    }
                }
                else if (comPinRow.PinName.ToUpper() == "SDO")
                {
                    adg1414Group.SDO = comPinRow.NetName;
                }
            }

            ADG1414Data.Sort(new SDOComparer());
            return ADG1414Data;
        }

        private string ConvertDesignName(string orgname)
        {
            //S0_U3901
            orgname = _regSNamePat1.Replace(orgname, "");
            return orgname;
        }

        private string ConvertSName(string orgSname)
        {
            //S0_BUCK0_LX4_UP1600_S1 BUCK0_LX4_UP1600_S1
            orgSname = _regSNamePat1.Replace(orgSname, "");
            //8.VI80_SEN16_S3 I80_SEN16_S3
            orgSname = _regSNamePat2.Replace(orgSname, "");
            return orgSname;
        }

        private class SDOComparer : IComparer<ADG1414Group>
        {
            Regex _DINNum = new Regex(@"(_DIN)(?<num>\d+)$", RegexOptions.IgnoreCase);
            public int Compare(ADG1414Group s, ADG1414Group t)
            {
                int sNum = -1;
                int tNum = -1;
                Match match = _DINNum.Match(s.SDO);
                if (match.Success)
                {
                    sNum = int.Parse(match.Groups["num"].Value);
                }
                else
                {
                    return 1;
                }

                match = _DINNum.Match(t.SDO);
                if (match.Success)
                {
                    tNum = int.Parse(match.Groups["num"].Value);
                }
                else
                {
                    return -1;
                }

                return sNum - tNum;
            }
        }
    }
}
