using System;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.DividerManager.Common
{
    public class DividerCommonLogic
    {
        public static void RemoveIgnoredSequence(HardIpPattern pattern)
        {
            var seqListInMisc = SearchInfo.GetSeqLstFromMiscInfo(pattern.MiscInfo);
            if (seqListInMisc == null)
                return;
            for (var i = 0; i < seqListInMisc.Count; i++)
            {
                var seqIndex = i + 1;
                if (seqListInMisc[i].Equals("N", StringComparison.OrdinalIgnoreCase))
                    pattern.MeasPins.RemoveAll(pin => pin.SequenceIndex == seqIndex);
            }
        }
    }
}