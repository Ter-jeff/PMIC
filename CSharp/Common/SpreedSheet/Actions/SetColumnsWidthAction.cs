#define WPF
using System.Collections.Generic;
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action for adjusting columns width.
    /// </summary>
    public class SetColumnsWidthAction : WorksheetReusableAction
    {
        private readonly Dictionary<int, ushort> backupCols = new Dictionary<int, ushort>();

        /// <summary>
        ///     Create instance for SetColsWidthAction
        /// </summary>
        /// <param name="col">Index of column start to set</param>
        /// <param name="count">Number of columns to be set</param>
        /// <param name="width">Width of column to be set</param>
        public SetColumnsWidthAction(int col, int count, ushort width)
            : base(new RangePosition(0, col, -1, count))
        {
            Width = width;
        }

        /// <summary>
        ///     Width to be set
        /// </summary>
        public ushort Width { get; set; }

        /// <summary>
        ///     Do this action
        /// </summary>
        public override void Do()
        {
            var col = Range.Col;
            var count = Range.Cols;

            backupCols.Clear();

            var c2 = col + count;
            for (var c = col; c < c2; c++)
            {
                var colHead = Worksheet.RetrieveColumnHeader(c);
                backupCols.Add(c, colHead.InnerWidth);
            }

            Worksheet.SetColumnsWidth(col, count, Width);
        }

        /// <summary>
        ///     Undo this action
        /// </summary>
        public override void Undo()
        {
            var col = Range.Col;
            var count = Range.Cols;

            Worksheet.SetColumnsWidth(col, count, c => backupCols[c]);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Set Cols Width: " + Width;
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new SetColumnsWidthAction(range.Col, range.Cols, Width);
        }
    }
}