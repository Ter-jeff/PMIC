using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ShmooLog.Base
{
    public class BinCutDataForShmoo
    {
        private string _binCutFileName;


        public List<ShmooBinCutFlowItem> ShmooBinCutFlowItems = new List<ShmooBinCutFlowItem>();

        public string BinCutFileName
        {
            get { return _binCutFileName; }
            set
            {
                _binCutFileName = value;
                var lVerPattern = @"Bin_*Cut.*(?<Version>V\dP\d)";
                var match = Regex.Match(value, lVerPattern, RegexOptions.IgnoreCase);
                if (match.Success)
                    BinCutVersion = match.Groups["Version"].ToString();
                else
                    BinCutVersion = "N/A";
            }
        }

        public string BinCutVersion { get; private set; }

        public string Job { get; set; }
        //public List<ShmooBinnedDomainItem> ShmonBinCutBinnedDomainItems = new List<ShmooBinnedDomainItem>();
    }

    public class ShmooBinCutFlowItem
    {
        public Dictionary<string, DomainPmodeRow> BinnedDomainPmodeData;
        public string Domain;

        public List<FlowPmodeRow> FlowModes;


        public ShmooBinCutFlowItem(string domain)
        {
            Domain = domain;
            FlowModes = new List<FlowPmodeRow>();
            BinnedDomainPmodeData = new Dictionary<string, DomainPmodeRow>();
        }
    }

    public class DomainPmodeRow
    {
        public double CPVmax;
        public double CPVmin;
        public bool IsOtherRail = false;
    }

    public class FlowPmodeRow
    {
        public string AtpgPmode;
        public string MbistPmode;
        public string OriginalPmode;
    }
}