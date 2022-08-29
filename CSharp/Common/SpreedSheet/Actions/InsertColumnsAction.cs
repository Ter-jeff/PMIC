#define WPF

using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Insert columns action
    /// </summary>
    public class InsertColumnsAction : WorksheetReusableAction
    {
        private int insertedCol = -1;

        /// <summary>
        ///     Create instance for InsertColumnsAction
        /// </summary>
        /// <param name="column">Index of column to insert</param>
        /// <param name="count">Number of columns to be insertted</param>
        public InsertColumnsAction(int column, int count)
            : base(RangePosition.Empty)
        {
            Column = column;
            Count = count;
        }

        /// <summary>
        ///     Index of column to insert new columns. Set to Control.ColCount to
        ///     append columns at end of columns.
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        ///     Number of columns to be inserted
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        ///     Do this action
        /// </summary>
        public override void Do()
        {
            insertedCol = Column;
            Worksheet.InsertColumns(Column, Count);
            Range = new RangePosition(0, Column, Worksheet.RowCount, Count);
        }

        /// <summary>
        ///     Undo this action
        /// </summary>
        public override void Undo()
        {
            Worksheet.DeleteColumns(insertedCol, Count);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Insert Columns";
        }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            return new InsertColumnsAction(range.Col, range.Cols);
        }
    }
}