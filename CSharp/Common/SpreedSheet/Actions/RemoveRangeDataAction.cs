#define WPF

using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action to remove data from specified range.
    /// </summary>
    public class RemoveRangeDataAction : WorksheetReusableAction
    {
        private object[,] backupData;

        /// <summary>
        ///     Create action to remove data from specified range.
        /// </summary>
        /// <param name="range">data from cells in this range will be removed.</param>
        public RemoveRangeDataAction(RangePosition range)
            : base(range)
        {
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new RemoveRangeDataAction(range);
        }

        /// <summary>
        ///     Do action to remove data from specified range.
        /// </summary>
        public override void Do()
        {
            backupData = Worksheet.GetRangeData(Range);
            Worksheet.DeleteRangeData(Range, true);
        }

        /// <summary>
        ///     Undo action to restore removed data.
        /// </summary>
        public override void Undo()
        {
            Worksheet.SetRangeData(Range, backupData);
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>friendly name of this action.</returns>
        public override string GetName()
        {
            return "Remove Cells Data";
        }
    }
}