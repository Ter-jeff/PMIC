#define WPF

using System.Diagnostics;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;

namespace unvell.ReoGrid.Actions
{
    internal class CutRangeAction : WorksheetReusableAction
    {
        private PartialGrid backupData;

        public CutRangeAction(RangePosition range, PartialGrid data) : base(range)
        {
            backupData = data;
        }

        public override void Do()
        {
            backupData =
                Worksheet.GetPartialGrid(Range, PartialGridCopyFlag.All, ExPartialGridCopyFlag.BorderOutsideOwner);
            Debug.Assert(backupData != null);

            Worksheet.DeleteRangeData(Range, true);
            Worksheet.RemoveRangeStyles(Range, PlainStyleFlag.All);
            Worksheet.RemoveRangeBorders(Range, BorderPositions.All);
        }

        public override void Undo()
        {
            Debug.Assert(backupData != null);
            Worksheet.SetPartialGrid(Range, backupData, PartialGridCopyFlag.All,
                ExPartialGridCopyFlag.BorderOutsideOwner);
        }

        public override string GetName()
        {
            return "Cut Range";
        }

        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new CutRangeAction(range, backupData);
        }
    }
}