using IgxlData.IgxlBase;
using PmicAutomation.Utility.PA.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.PA.Base
{

    public class DgsPool
    {
        private PaSheet _DGSReferenceSheet = null;

        private readonly int _siteCnt;
        private const string ConDcDiffMeter = "DCDiffMeter";
        private const string ConNC = "N/C";

        private Dictionary<int, List<PaRow>> _siteParowDic;

        public DgsPool(PaSheet referenceSheet, int siteCnt)
        {
            _DGSReferenceSheet = referenceSheet;
            _siteCnt = siteCnt;
            _siteParowDic = new Dictionary<int, List<PaRow>>();
        }

        public List<ChannelMapRow> GenAllDgsChannelMap()
        {
            List<ChannelMapRow> allChanMapRows = new List<ChannelMapRow>();

            StringBuilder sb = new StringBuilder();
            this.FormatPrecheck(ref sb);

            allChanMapRows.AddRange(GenDgsChanMapRows("DC30", 4));
            allChanMapRows.AddRange(GenDgsChanMapRows("UVI80", 8));
            return allChanMapRows;
        }

        private List<ChannelMapRow> GenDgsChanMapRows(string type, int pinNum)
        {
            List<ChannelMapRow> chanMapRows = new List<ChannelMapRow>();
            List<PaRow> firstSiteRows = _siteParowDic.First().Value;
            Dictionary<string, int> assignmentDic = new Dictionary<string, int>();
            for (int i = 0; i < firstSiteRows.Count; i++)
            {
                if (firstSiteRows[i].BumpName.Contains(type) && !assignmentDic.ContainsKey(firstSiteRows[i].Assignment))
                {
                    assignmentDic.Add(firstSiteRows[i].Assignment, i);
                }
            }

            List<int> indexList = assignmentDic.Values.OrderBy(o => o).ToList();
            int elementCount = indexList.Count < pinNum ? pinNum : indexList.Count;
            for (int i = 0; i < elementCount; i++)
            {
                ChannelMapRow chanMapRow = new ChannelMapRow();
                if (i < indexList.Count)
                {
                    int index = indexList[i];
                    foreach (var siteParowItem in _siteParowDic)
                    {
                        chanMapRow.Sites.Add(siteParowItem.Value[index].Assignment);
                    }
                    chanMapRow.Type = ConDcDiffMeter;
                }
                else
                {
                    for (int j = 0; j < _siteCnt; j++)
                    {
                        chanMapRow.Sites.Add("");
                    }
                    chanMapRow.Type = ConNC;
                }

                chanMapRow.DiviceUnderTestPinName = type + "_DGS_" + i + "_DM";

                chanMapRows.Add(chanMapRow);
            }

            return chanMapRows;
        }


        public bool FormatPrecheck(ref StringBuilder errorMessage)
        {
            if (_DGSReferenceSheet == null)
                return false;

            errorMessage = new StringBuilder();
            List<PaRow> dgsRows = _DGSReferenceSheet.Rows;

            foreach (PaRow dgsRow in dgsRows)
            {
                int site = -1;
                bool flag = int.TryParse(dgsRow.Site, out site);
                if (flag && site >= 0)
                {
                    if (_siteParowDic.ContainsKey(site))
                    {
                        _siteParowDic[site].Add(dgsRow);
                    }
                    else
                    {
                        _siteParowDic.Add(site, new List<PaRow>());
                        _siteParowDic[site].Add(dgsRow);
                    }
                }
            }

            int sum = -1;
            bool result = true;
            _siteParowDic.Values.ToList().ForEach(rows =>
            {
                if (sum == -1)
                {
                    sum = rows.Count;
                }

                if (sum != rows.Count)
                {
                    result = false;
                }
            });

            if (!result)
                errorMessage.AppendLine("The format per site in DGS Reference is not matched.");

            if (_siteCnt != _siteParowDic.Keys.Count)
            {
                result = false;
                errorMessage.AppendLine("The site number of DGS Reference should be as same as PA File's site number.");
            }

            return result;
        }


        public int GetDgsPinCount(string type)
        {
            StringBuilder sb = new StringBuilder();
            this.FormatPrecheck(ref sb);

            List<PaRow> firstSiteRows = _siteParowDic.First().Value;
            Dictionary<string, int> assignmentDic = new Dictionary<string, int>();
            for (int i = 0; i < firstSiteRows.Count; i++)
            {
                if (firstSiteRows[i].BumpName.Contains(type) && !assignmentDic.ContainsKey(firstSiteRows[i].Assignment))
                {
                    assignmentDic.Add(firstSiteRows[i].Assignment, i);
                }
            }

            return assignmentDic.Count;
        }

    }
}
