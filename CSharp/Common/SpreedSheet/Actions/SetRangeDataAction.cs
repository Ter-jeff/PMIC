#define WPF
using System.Diagnostics;
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Create action to set data into specified range of spreadsheet.
    /// </summary>
    public class SetRangeDataAction : WorksheetReusableAction
    {
        private readonly object[,] data;
        private object[,] backupData;

        /// <summary>
        ///     Create action to set data into specified range of spreadsheet.
        /// </summary>
        /// <param name="range">range to set specified data.</param>
        /// <param name="data">data to be set.</param>
        public SetRangeDataAction(RangePosition range, object[,] data)
            : base(range)
        {
            this.data = data;
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetRangeDataAction(range, data);
        }

        /// <summary>
        ///     Do action to set data into specified range of spreadsheet
        /// </summary>
        public override void Do()
        {
            backupData = Worksheet.GetRangeData(Range);
            Debug.Assert(backupData != null);
            Worksheet.SetRangeData(Range, data, true);
            Worksheet.SelectRange(Range);
        }

        /// <summary>
        ///     Undo action to remove data which has been set into specified range of spreadsheet
        /// </summary>
        public override void Undo()
        {
            Debug.Assert(backupData != null);
            Worksheet.SetRangeData(Range, backupData);
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>friendly name of this action.</returns>
        public override string GetName()
        {
            return "Set Cells Data";
        }
    }
}