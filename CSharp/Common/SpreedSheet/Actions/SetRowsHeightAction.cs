#define WPF

using System.Collections.Generic;
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Set height of row action
    /// </summary>
    public class SetRowsHeightAction : WorksheetReusableAction
    {
        private readonly Dictionary<int, RowHeadData> backupRows = new Dictionary<int, RowHeadData>();

        /// <summary>
        ///     Create instance for SetRowsHeightAction
        /// </summary>
        /// <param name="row">Index of row start to set</param>
        /// <param name="count">Number of rows to be set</param>
        /// <param name="height">New height to set to specified rows</param>
        public SetRowsHeightAction(int row, int count, ushort height)
            : base(new RangePosition(row, 0, count, -1))
        {
            Height = height;
        }

        /// <summary>
        ///     Height to be set
        /// </summary>
        public ushort Height { get; set; }

        /// <summary>
        ///     Do this action
        /// </summary>
        public override void Do()
        {
            var row = Range.Row;
            var count = Range.Rows;

            backupRows.Clear();

            var r2 = row + count;
            for (var r = row; r < r2; r++)
            {
                var rowHead = Worksheet.RetrieveRowHeader(r);

                backupRows.Add(r, new RowHeadData
                {
                    autoHeight = rowHead.IsAutoHeight,
                    row = rowHead.Row,
                    height = rowHead.InnerHeight
                });

                // disable auto-height-adjusting if user has changed height of this row
                rowHead.IsAutoHeight = false;
            }

            Worksheet.SetRowsHeight(row, count, Height);
        }

        /// <summary>
        ///     Undo this action
        /// </summary>
        public override void Undo()
        {
            Worksheet.SetRowsHeight(Range.Row, Range.Rows, r => backupRows[r].height, true);

            foreach (var r in backupRows.Keys)
            {
                var rowHead = Worksheet.RetrieveRowHeader(r);
                rowHead.IsAutoHeight = backupRows[r].autoHeight;
            }
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Set Rows Height: " + Height;
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetRowsHeightAction(range.Row, range.Rows, Height);
        }

        internal struct RowHeadData
        {
            internal int row;
            internal ushort height;
            internal bool autoHeight;
        }
    }
}