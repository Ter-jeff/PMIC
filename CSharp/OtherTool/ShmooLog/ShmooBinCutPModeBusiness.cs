using ShmooLog.Base;
using System.Linq;
using System.Text.RegularExpressions;

namespace ShmooLog
{
    public static class ShmooBinCutPModeBusiness
    {
        public static DomainPmodeRow GetTargerBinnedDominItem(BinCutDataForShmoo binCutData, string domain,
            ref string pmode)
        {
            var flowItem = binCutData.ShmooBinCutFlowItems.FirstOrDefault(p => p.Domain.Equals(domain));
            if (flowItem == null) return null;


            if (domain.Equals("FIXED"))
            {
                pmode = "FIXED";
                if (flowItem.BinnedDomainPmodeData.ContainsKey(pmode))
                    return flowItem.BinnedDomainPmodeData[pmode];
            }
            else if (domain.Equals("LOW"))
            {
                pmode = "LOW";
                if (flowItem.BinnedDomainPmodeData.ContainsKey(pmode))
                    return flowItem.BinnedDomainPmodeData[pmode];
            }


            var targetMode = "";
            FlowPmodeRow flowPmodeRow;
            var searchMode = pmode;

            // search Oringinal 
            flowPmodeRow = flowItem.FlowModes.FirstOrDefault(i => i.OriginalPmode.Equals(searchMode));
            if (flowPmodeRow != null)
            {
                targetMode = flowPmodeRow.OriginalPmode;
                if (flowItem.BinnedDomainPmodeData.ContainsKey(targetMode))
                    return flowItem.BinnedDomainPmodeData[targetMode];
            }

            // search ATPG
            flowPmodeRow = flowItem.FlowModes.FirstOrDefault(i => i.AtpgPmode.Equals(searchMode));
            if (flowPmodeRow != null)
            {
                targetMode = flowPmodeRow.OriginalPmode;
                if (flowItem.BinnedDomainPmodeData.ContainsKey(targetMode))
                    return flowItem.BinnedDomainPmodeData[targetMode];
            }

            // search Mbist
            flowPmodeRow = flowItem.FlowModes.FirstOrDefault(i => i.MbistPmode.Equals(searchMode));
            if (flowPmodeRow != null)
            {
                targetMode = flowPmodeRow.OriginalPmode;
                if (flowItem.BinnedDomainPmodeData.ContainsKey(targetMode))
                    return flowItem.BinnedDomainPmodeData[targetMode];
            }

            return null;
        }

        public static string DomainGetting(string binCutDomain)
        {
            string domain;
            const string lStrMatchPattern = @"VDD(?<str>\w+)";

            switch (binCutDomain)
            {
                case "VDDFIXEDGROUP":
                    domain = "FIXED";
                    break;
                case "VDDLOWGROUP":
                    domain = "LOW";
                    break;
                default:
                    domain = Regex.Match(binCutDomain, lStrMatchPattern, RegexOptions.IgnoreCase).Groups["str"]
                        .ToString();
                    break;
            }

            return domain;
        }
    }
}