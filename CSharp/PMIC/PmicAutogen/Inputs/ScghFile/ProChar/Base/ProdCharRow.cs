using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Base
{
    public class ProdCharRow
    {
        private List<PatternWithMode> _payloadList;

        #region result

        public InstanceRow InstanceRow;

        #endregion

        public bool SkipCheckRule;

        public ProdCharRow(IProdCharSheetRow prodCharRow)
        {
            ProdCharItem = prodCharRow;
            InitList = new Dictionary<int, PatternWithMode>();
            PayloadList = new List<PatternWithMode>();
            InitAliasList = new List<string>();
            PayloadAliasList = new List<string>();
        }

        public int RowNum { set; get; }
        public string Prefix { set; get; }
        public string InitPatSetNameByNamingRule { set; get; }
        public string PatSetName { set; get; }
        public string PayLoadName { get; set; }
        public string InstanceName { set; get; }

        public List<PatternWithMode> PayloadList
        {
            set
            {
                if (value.Count > 1)
                {
                    var patSet = new PatSet();
                    PayLoadName = patSet.GetNewPatSetName(value.Select(x => x.PatternName).ToList());
                }
                else if (value.Count == 1)
                {
                    PayLoadName = value[0].PatternName;
                }
                else
                {
                    PayLoadName = "";
                }

                _payloadList = value;
            }
            get { return _payloadList; }
        }

        public List<string> PayloadAliasList { set; get; }
        public Dictionary<int, PatternWithMode> InitList { set; get; }
        public List<string> InitAliasList { set; get; }
        public bool InitPatternMissing { set; get; }
        public IProdCharSheetRow ProdCharItem { set; get; }
        public bool Nop { set; get; }
        public NopType NopType { set; get; }
        public string PayloadType { set; get; }

        //public override int GetHashCode()
        //{
        //    int res = 0x2D2816FE;
        //    foreach (var item in InitList)
        //        res = res * 31 + item.Value.PatternName.GetHashCode();
        //    foreach (var item in PayloadList)
        //        res = res * 31 + item.PatternName.GetHashCode();
        //    return res;
        //}

        public List<List<string>> GetTrackingGroup(string supplyVoltage)
        {
            //2017/4/7: Since M8P change there supply voltage gen CZ rule, 
            //ex: prodCharTestInstance.SupplyVoltage = "(A),B,C,(D,E),(F,G)",
            //    In tracking mode will generate A<-B<-C, D<-E, F<-G 3 CZ setup
            // C, A_T_B will generate  C tracking A,B   and A tracking B
            var allTrackingGroups = new List<List<string>>();
            var inputStr = supplyVoltage.Replace(" ", "");
            var matches = Regex.Matches(inputStr, @"[(](?<AAA>[\w]*\,[\w|,]+)[)]", RegexOptions.IgnoreCase);
            foreach (var match in matches)
            {
                var strTmp = match.ToString();
                allTrackingGroups.Add(Regex.Replace(strTmp, @"[()]", "")
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList());
                inputStr = inputStr.Replace(match.ToString(), "");
            }

            if (inputStr != "")
            {
                var grpTmp = Regex.Replace(inputStr, @"[()]", "")
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                //Split using _T_
                for (var i = 0; i < grpTmp.Count; i++)
                {
                    if (!Regex.IsMatch(grpTmp[i], "_T_", RegexOptions.IgnoreCase))
                        continue;
                    var groupFromT = Regex.Split(grpTmp[i], "_T_", RegexOptions.IgnoreCase).ToList();
                    if (!allTrackingGroups.Exists(p => GroupEqual(p, groupFromT)))
                        allTrackingGroups.Add(groupFromT);
                    grpTmp.RemoveAt(i);

                    grpTmp.InsertRange(i, groupFromT);
                    i += groupFromT.Count;
                    i--;
                }

                if (grpTmp.Count > 0)
                    allTrackingGroups.Add(grpTmp);
            }

            return allTrackingGroups;
        }

        public List<string> GetSinglePins(string supplyVoltage)
        {
            var groups = GetTrackingGroup(supplyVoltage);
            var pins = new List<string>();
            foreach (var group in groups)
                foreach (var pin in group)
                    if (!pins.Contains(pin))
                        pins.Add(pin);

            return pins;
        }

        private bool GroupEqual(List<string> group1, List<string> group2)
        {
            return !group1.Except(group2).Any() && !group2.Except(group1).Any();
        }

        public string Get2DCharNamePeriod(string pinName, string block)
        {
            return "Shmoo_2D_" + pinName + "_vs_" + block + "_PERIOD";
        }

        public string Get1DCharNamePeriod(string pinName)
        {
            return "Shmoo_1D_" + pinName;
        }

        public ProdCharRowScan NewProdCharRowScan()
        {
            var prodCharRowScan = new ProdCharRowScan(ProdCharItem);
            prodCharRowScan.InitList = InitList;
            prodCharRowScan.PayloadList = PayloadList.ToList();
            prodCharRowScan.InstanceName = InstanceName;
            prodCharRowScan.PatSetName = PatSetName;
            prodCharRowScan.InitPatSetNameByNamingRule = InitPatSetNameByNamingRule;
            prodCharRowScan.InitPatternMissing = InitPatternMissing;
            prodCharRowScan.ProdCharItem = ProdCharItem;
            prodCharRowScan.SkipCheckRule = SkipCheckRule;
            return prodCharRowScan;
        }

        public ProdCharRowMbist NewProdCharRowMbist()
        {
            var prodCharRowMbist = new ProdCharRowMbist(ProdCharItem);
            prodCharRowMbist.InitList = InitList;
            prodCharRowMbist.PayloadList = PayloadList.ToList();
            prodCharRowMbist.InstanceName = InstanceName;
            prodCharRowMbist.PatSetName = PatSetName;
            prodCharRowMbist.InitPatSetNameByNamingRule = InitPatSetNameByNamingRule;
            prodCharRowMbist.InitPatternMissing = InitPatternMissing;
            prodCharRowMbist.ProdCharItem = ProdCharItem;
            prodCharRowMbist.SkipCheckRule = SkipCheckRule;
            return prodCharRowMbist;
        }
    }
}