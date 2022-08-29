#define WPF

using SpreedSheet.Core;
using unvell.ReoGrid.DataFormat;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Set range data format action.
    /// </summary>
    public class SetRangeDataFormatAction : WorksheetReusableAction
    {
        private readonly CellDataFormatFlag format;
        private readonly object formatArgs;
        private PartialGrid backupData;

        /// <summary>
        ///     Create instance for SetRangeDataFormatAction.
        /// </summary>
        /// <param name="range">Range to be appiled this action.</param>
        /// <param name="format">Format type of cell to be set.</param>
        /// <param name="dataFormatArgs">Argument belongs to format type to be set.</param>
        public SetRangeDataFormatAction(RangePosition range, CellDataFormatFlag format,
            object dataFormatArgs)
            : base(range)
        {
            this.format = format;
            formatArgs = dataFormatArgs;
        }

        /// <summary>
        ///     Do this operation.
        /// </summary>
        public override void Do()
        {
            backupData = Worksheet.GetPartialGrid(Range, PartialGridCopyFlag.CellData, ExPartialGridCopyFlag.None);
            Worksheet.SetRangeDataFormat(Range, format, formatArgs);
        }

        /// <summary>
        ///     Undo this operation.
        /// </summary>
        public override void Undo()
        {
            Worksheet.SetPartialGrid(Range, backupData);
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetRangeDataFormatAction(range, format, formatArgs);
        }

        /// <summary>
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>friendly name of this action.</returns>
        public override string GetName()
        {
            return "Set Cells Format: " + format;
        }
    }
}