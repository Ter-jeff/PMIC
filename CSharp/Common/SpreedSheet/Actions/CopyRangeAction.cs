#define WPF

using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action to copy the specified range from a position to another position.
    /// </summary>
    public class CopyRangeAction : BaseWorksheetAction
    {
        private PartialGrid backupGrid;

        /// <summary>
        ///     Construct this action to move specified range from a position to another position
        /// </summary>
        /// <param name="fromRange">range to be moved</param>
        /// <param name="toPosition">position to be moved to</param>
        public CopyRangeAction(RangePosition fromRange, CellPosition toPosition)
        {
            ContentFlags = PartialGridCopyFlag.All;
            FromRange = fromRange;
            ToPosition = toPosition;
        }

        /// <summary>
        ///     Specifies the content to be moved: data, borders and styles.
        /// </summary>
        public PartialGridCopyFlag ContentFlags { get; set; }

        /// <summary>
        ///     Range to be moved
        /// </summary>
        public RangePosition FromRange { get; set; }

        /// <summary>
        ///     Position that range will be moved to
        /// </summary>
        public CellPosition ToPosition { get; set; }

        /// <summary>
        ///     Do this action.
        /// </summary>
        public override void Do()
        {
            var targetRange = new RangePosition(
                ToPosition.Row, ToPosition.Col,
                FromRange.Rows, FromRange.Cols);

            backupGrid = Worksheet.GetPartialGrid(targetRange);

            Worksheet.CopyRange(FromRange, targetRange);

            Worksheet.SelectionRange = targetRange;
        }

        /// <summary>
        ///     Undo this action.
        /// </summary>
        public override void Undo()
        {
            var targetRange = new RangePosition(
                ToPosition.Row, ToPosition.Col,
                FromRange.Rows, FromRange.Cols);

            Worksheet.SetPartialGrid(targetRange, backupGrid);

            Worksheet.SelectionRange = FromRange;
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>friendly name of this action.</returns>
        public override string GetName()
        {
            return "Copy Range";
        }
    }
}