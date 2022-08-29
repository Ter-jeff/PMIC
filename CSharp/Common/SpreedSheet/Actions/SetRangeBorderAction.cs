#define WPF
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action to set borders to specified range
    /// </summary>
    public class SetRangeBorderAction : WorksheetReusableAction
    {
        private PartialGrid backupData;

        /// <summary>
        ///     Create action that perform setting border to a range
        /// </summary>
        /// <param name="range">Range to be appiled this action</param>
        /// <param name="pos">Position of range to set border</param>
        /// <param name="styles">Style of border</param>
        public SetRangeBorderAction(RangePosition range, BorderPositions pos, RangeBorderStyle styles)
            : this(range, new[] { new RangeBorderInfo(pos, styles) })
        {
        }

        /// <summary>
        ///     Create action that perform setting border to a range
        /// </summary>
        /// <param name="range">Range to be appiled this action</param>
        /// <param name="styles">Style of border</param>
        public SetRangeBorderAction(RangePosition range, RangeBorderInfo[] styles)
            : base(range)
        {
            Borders = styles;
        }

        /// <summary>
        ///     Borders to be set
        /// </summary>
        public RangeBorderInfo[] Borders { get; set; }

        /// <summary>
        ///     Do this action
        /// </summary>
        public override void Do()
        {
            backupData = Worksheet.GetPartialGrid(Range, PartialGridCopyFlag.BorderAll,
                ExPartialGridCopyFlag.BorderOutsideOwner);

            for (var i = 0; i < Borders.Length; i++) Worksheet.SetRangeBorders(Range, Borders[i].Pos, Borders[i].Style);
        }

        /// <summary>
        ///     Undo this action
        /// </summary>
        public override void Undo()
        {
            Worksheet.SetPartialGrid(Range, backupData, PartialGridCopyFlag.BorderAll,
                ExPartialGridCopyFlag.BorderOutsideOwner);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Set Range Border";
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetRangeBorderAction(range, Borders);
        }
    }
}