#define WPF

using System.Diagnostics;
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action to set partial grid.
    /// </summary>
    public class SetPartialGridAction : WorksheetReusableAction
    {
        private readonly PartialGrid data;
        private PartialGrid backupData;

        /// <summary>
        ///     Create action to set partial grid.
        /// </summary>
        /// <param name="range">target range to set partial grid.</param>
        /// <param name="data">partial grid to be set.</param>
        public SetPartialGridAction(RangePosition range, PartialGrid data)
            : base(range)
        {
            this.data = data;
        }

        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetPartialGridAction(range, data);
        }

        /// <summary>
        ///     Do action to set partial grid.
        /// </summary>
        public override void Do()
        {
            backupData =
                Worksheet.GetPartialGrid(Range, PartialGridCopyFlag.All, ExPartialGridCopyFlag.BorderOutsideOwner);
            Debug.Assert(backupData != null);
            Range = Worksheet.SetPartialGridRepeatly(Range, data);
            Worksheet.SelectRange(Range);
        }

        /// <summary>
        ///     Undo action to restore setting partial grid.
        /// </summary>
        public override void Undo()
        {
            Debug.Assert(backupData != null);
            Worksheet.SetPartialGrid(Range, backupData, PartialGridCopyFlag.All,
                ExPartialGridCopyFlag.BorderOutsideOwner);
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>Friendly name of this action.</returns>
        public override string GetName()
        {
            return "Set Partial Grid";
        }
    }
}