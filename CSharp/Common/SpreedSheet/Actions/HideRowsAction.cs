#define WPF

using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Hide specified rows action
    /// </summary>
    public class HideRowsAction : WorksheetReusableAction
    {
        /// <summary>
        ///     Create action to hide specified rows.
        /// </summary>
        /// <param name="row">Zero-based row index to start hiding.</param>
        /// <param name="count">Number of rows to be hidden.</param>
        public HideRowsAction(int row, int count)
            : base(new RangePosition(row, 0, count, -1))
        {
        }

        /// <summary>
        ///     Do action to hide specified rows.
        /// </summary>
        public override void Do()
        {
            Worksheet.HideRows(Range.Row, Range.Rows);
        }

        /// <summary>
        ///     Undo action to show hidden rows.
        /// </summary>
        public override void Undo()
        {
            Worksheet.ShowRows(Range.Row, Range.Rows);
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new HideRowsAction(range.Row, range.Rows);
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>friendly name of this action.</returns>
        public override string GetName()
        {
            return "Hide Rows";
        }
    }
}