using System;
using SpreedSheet.Core;
using unvell.ReoGrid;

namespace SpreedSheet.Interface
{
    internal interface IRangePickableControl
    {
        void PickRange(Func<Worksheet, RangePosition, bool> handler);
        void EndPickRange();
        void StartPickRangeAndCopyStyle();
    }
}