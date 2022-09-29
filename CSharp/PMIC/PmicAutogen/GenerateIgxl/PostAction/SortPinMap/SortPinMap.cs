using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.PostAction.SortPinMap
{
    public class SortPinMap
    {
        private const string TimeDomain = "TimeDomain";

        public void Sort(PinMapSheet pPinMapSheet)
        {
            if (pPinMapSheet == null) return;
            var lTimeDomainList = new List<PinGroup>();
            for (var i = 0; i < pPinMapSheet.GroupList.Count; i++)
                if (pPinMapSheet.GroupList[i].PinType.Equals(TimeDomain, StringComparison.OrdinalIgnoreCase))
                {
                    lTimeDomainList.Add(pPinMapSheet.GroupList[i]);
                    pPinMapSheet.GroupList.RemoveAt(i);
                    i--;
                }

            pPinMapSheet.GroupList.AddRange(lTimeDomainList);
        }
    }
}