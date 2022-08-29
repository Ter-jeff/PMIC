#define WPF

using SpreedSheet.Core;
using SpreedSheet.Core.Enum;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Remove style from specified range action
    /// </summary>
    public class RemoveRangeStyleAction : WorksheetReusableAction
    {
        private PartialGrid backupData;

        /// <summary>
        ///     Create instance for action to remove style from specified range.
        /// </summary>
        /// <param name="range">Styles from this specified range to be removed</param>
        /// <param name="flag">Style flag indicates what type of style should be removed</param>
        public RemoveRangeStyleAction(RangePosition range, PlainStyleFlag flag)
            : base(range)
        {
            Flag = flag;
        }

        /// <summary>
        ///     Style flag indicates what type of style to be handled.
        /// </summary>
        public PlainStyleFlag Flag { get; set; }

        /// <summary>
        ///     Do this action
        /// </summary>
        public override void Do()
        {
            backupData = Worksheet.GetPartialGrid(Range);
            Worksheet.RemoveRangeStyles(Range, Flag);
        }

        /// <summary>
        ///     Undo this action
        /// </summary>
        public override void Undo()
        {
            Worksheet.SetPartialGrid(Range, backupData, PartialGridCopyFlag.CellStyle);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Delete Style";
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new RemoveRangeStyleAction(Range, Flag);
        }
    }
}