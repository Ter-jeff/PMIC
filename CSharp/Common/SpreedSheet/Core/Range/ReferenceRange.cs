#define WPF

using System;
using SpreedSheet.Core;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Represents a range object refer to spreadsheet
    /// </summary>
    public class ReferenceRange : IRange
    {
        #region Value & Properties

        /// <summary>
        ///     Get or set the worksheet which contains this range
        /// </summary>
        public Worksheet Worksheet { get; internal set; }

        private Cell startCell;

        private Cell endCell;
        //private ReoGridRange range = ReoGridRange.Empty;

        /// <summary>
        ///     Get or set start position.
        /// </summary>
        public CellPosition StartPos
        {
            get { return startCell.Position; }
            set { startCell = Worksheet.CreateAndGetCell(Worksheet.FixPos(value)); }
        }

        /// <summary>
        ///     Get or set end position.
        /// </summary>
        public CellPosition EndPos
        {
            get { return endCell.Position; }
            set { endCell = Worksheet.CreateAndGetCell(Worksheet.FixPos(value)); }
        }

        /// <summary>
        ///     Zero-based number of row to locate the start position of this range.
        /// </summary>
        public int Row
        {
            get { return startCell.Row; }
            set { startCell = Worksheet.CreateAndGetCell(value, startCell.Column); }
        }

        /// <summary>
        ///     Zero-based number of column to locate the start position of this range.
        /// </summary>
        public int Col
        {
            get { return startCell.Column; }
            set { startCell = Worksheet.CreateAndGetCell(startCell.Row, value); }
        }

        /// <summary>
        ///     Get or set number of rows.
        /// </summary>
        public int Rows
        {
            get { return Position.Rows; }
            set { endCell = Worksheet.CreateAndGetCell(startCell.Row + value, endCell.Column); }
        }

        /// <summary>
        ///     Get or set number of columns.
        /// </summary>
        public int Cols
        {
            get { return Position.Cols; }
            set { endCell = Worksheet.CreateAndGetCell(endCell.Row, startCell.Column + value); }
        }

        /// <summary>
        ///     Get or set end number of row.
        /// </summary>
        public int EndRow
        {
            get { return endCell.Row; }
            set { endCell = Worksheet.CreateAndGetCell(new CellPosition(value, endCell.Column)); }
        }

        /// <summary>
        ///     Get or set end number of column.
        /// </summary>
        public int EndCol
        {
            get { return endCell.Column; }
            set { endCell = Worksheet.CreateAndGetCell(endCell.Row, value); }
        }

        /// <summary>
        ///     Get or set the position of range on worksheet.
        /// </summary>
        public RangePosition Position
        {
            get { return new RangePosition(startCell.Position, endCell.Position); }
            set
            {
                var range = Worksheet.FixRange(value);

                startCell = Worksheet.CreateAndGetCell(range.StartPos);
                endCell = Worksheet.CreateAndGetCell(range.EndPos);
            }
        }

        #region Constructors

        internal ReferenceRange(Worksheet worksheet, Cell startCell, Cell endCell)
        {
            if (worksheet == null)
                throw new ArgumentNullException("worksheet", "cannot create refereced range with null worksheet");

            Worksheet = worksheet;
            this.startCell = startCell;
            this.endCell = endCell;
        }

        internal ReferenceRange(Worksheet worksheet, CellPosition startPos, CellPosition endPos)
            : this(worksheet, worksheet.CreateAndGetCell(startPos), worksheet.CreateAndGetCell(endPos))
        {
        }

        internal ReferenceRange(Worksheet worksheet, string address)
            : this(worksheet, new RangePosition(address))
        {
            // construct from address identifier
        }

        internal ReferenceRange(Worksheet worksheet, RangePosition range)
            : this(worksheet, worksheet.CreateAndGetCell(range.StartPos), worksheet.CreateAndGetCell(range.EndPos))
        {
            // construct from range position
        }

        internal ReferenceRange(Worksheet worksheet, CellPosition pos)
            : this(worksheet, pos, pos)
        {
            // construct from single cell position
        }

        #endregion // Constructors

        #endregion Value & Properties

        #region Utility

        /// <summary>
        ///     Check whether or not the specified position is contained by this range.
        /// </summary>
        /// <param name="pos">Position to be checked.</param>
        /// <returns>True if specified position is contained by this range.</returns>
        public bool Contains(CellPosition pos)
        {
            var startPos = StartPos;
            var endPos = EndPos;

            return pos.Row >= startPos.Row && pos.Col >= startPos.Col
                                           && pos.Row <= endPos.Row && pos.Col <= endPos.Col;
        }

        /// <summary>
        ///     Check whether or not a specified range is contained by this range.
        /// </summary>
        /// <param name="range">Range position to be checked.</param>
        /// <returns>True if the specified range is contained by this range; Otherwise return false.</returns>
        public bool Contains(ReferenceRange range)
        {
            return startCell.InternalRow <= range.startCell.InternalRow
                   && startCell.InternalCol <= range.startCell.InternalCol
                   && endCell.InternalRow >= range.endCell.InternalRow
                   && endCell.InternalCol >= range.endCell.InternalCol;
        }

        /// <summary>
        ///     Check whether or not a specified range is contained by this range.
        /// </summary>
        /// <param name="range">Range position to be checked.</param>
        /// <returns>True if the specified range is contained by this range; Otherwise return false.</returns>
        public bool Contains(RangePosition range)
        {
            return startCell.InternalRow <= range.Row
                   && startCell.InternalCol <= range.Col
                   && endCell.InternalRow >= range.EndRow
                   && endCell.InternalCol >= range.EndCol;
        }

        /// <summary>
        ///     Check whether or not that the specified range intersects with this range.
        /// </summary>
        /// <param name="range">The range to be checked.</param>
        /// <returns>True if specified range intersects with this range.</returns>
        public bool IntersectWith(RangePosition range)
        {
            return Position.IntersectWith(range);
        }

        /// <summary>
        ///     Check whether or not that the specified range intersects with this range.
        /// </summary>
        /// <param name="range">The range to be checked.</param>
        /// <returns>True if specified range intersects with this range.</returns>
        public bool IntersectWith(ReferenceRange range)
        {
            return IntersectWith(range.Position);
        }

        /// <summary>
        ///     Convert to ReoGridRange structure.
        /// </summary>
        /// <param name="refRange">The object to be converted.</param>
        /// <returns>ReoGridRange structure converted from reference range instance.</returns>
        public static implicit operator RangePosition(ReferenceRange refRange)
        {
            return refRange.Position;
        }

        /// <summary>
        ///     Convert reference range into description string.
        /// </summary>
        /// <returns>String to describe this reference range.</returns>
        public override string ToString()
        {
            return Position.ToString();
        }

        /// <summary>
        ///     Convert referenced range into address position string.
        /// </summary>
        /// <returns>Address position string to describe this range on worksheet.</returns>
        public virtual string ToAddress()
        {
            return Position.ToAddress();
        }

        /// <summary>
        ///     Convert referenced range into absolute address position string.
        /// </summary>
        /// <returns>Absolute address position string to describe this range on worksheet.</returns>
        public virtual string ToAbsoluteAddress()
        {
            return Position.ToAbsoluteAddress();
        }

        #endregion // Utility

        #region Control API Routines

        /// <summary>
        ///     Get or set data of this range.
        /// </summary>
        public object Data
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeData(this);
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeData(this, value);
            }
        }

        /// <summary>
        ///     Select this range.
        /// </summary>
        public void Select()
        {
            CheckForOwnerAssociated();

            Worksheet.SelectRange(Position);
        }

        #region Style Wrapper

        private ReferenceRangeStyle referenceStyle;

        /// <summary>
        ///     Get the style set from this range.
        /// </summary>
        public ReferenceRangeStyle Style
        {
            get
            {
                CheckForOwnerAssociated();

                if (referenceStyle == null) referenceStyle = new ReferenceRangeStyle(Worksheet, this);

                return referenceStyle;
            }
        }

        #endregion // Style Wrapper

        #region Border Wraps

        private RangeBorderProperty borderProperty;

        public RangeBorderProperty Border
        {
            get
            {
                if (borderProperty == null) borderProperty = new RangeBorderProperty(this);

                return borderProperty;
            }
        }

        /// <summary>
        ///     Get or set left border styles for range.
        /// </summary>
        public RangeBorderStyle BorderLeft
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.Left).Left;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.Left, value);
            }
        }

        /// <summary>
        ///     Get or set top border styles for range.
        /// </summary>
        public RangeBorderStyle BorderTop
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.Top).Top;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.Top, value);
            }
        }

        /// <summary>
        ///     Get or set right border styles for range.
        /// </summary>
        public RangeBorderStyle BorderRight
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.Right).Right;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.Right, value);
            }
        }

        /// <summary>
        ///     Get or set bottom border styles for range.
        /// </summary>
        public RangeBorderStyle BorderBottom
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.Bottom).Bottom;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.Bottom, value);
            }
        }

        /// <summary>
        ///     Get or set all inside borders style for range.
        /// </summary>
        public RangeBorderStyle BorderInsideAll
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.InsideAll)
                    .InsideHorizontal; // TODO: no outline available here
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.InsideAll, value);
            }
        }

        /// <summary>
        ///     Get or set all horizontal border styles for range.
        /// </summary>
        public RangeBorderStyle BorderInsideHorizontal
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.InsideHorizontal).InsideHorizontal;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.InsideHorizontal, value);
            }
        }

        /// <summary>
        ///     Get or set all vertical border styles for range.
        /// </summary>
        public RangeBorderStyle BorderInsideVertical
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position, BorderPositions.InsideVertical).InsideVertical;
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.InsideVertical, value);
            }
        }

        /// <summary>
        ///     Get or set all outside border styles for range.
        /// </summary>
        public RangeBorderStyle BorderOutside
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet == null
                    ? RangeBorderStyle.Empty
                    : Worksheet.GetRangeBorders(Position, BorderPositions.Outside)
                        .Left; // TODO: no outline available here
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.Outside, value);
            }
        }

        /// <summary>
        ///     Get or set all inside border styles for range.
        /// </summary>
        public RangeBorderStyle BorderAll
        {
            get
            {
                CheckForOwnerAssociated();

                return Worksheet.GetRangeBorders(Position).Left; // TODO: no outline available here
            }
            set
            {
                CheckForOwnerAssociated();

                Worksheet.SetRangeBorders(Position, BorderPositions.All, value);
            }
        }

        #endregion // Border Wraps

        #region Merge & Group

        /// <summary>
        ///     Merge this range into single cell
        /// </summary>
        public void Merge()
        {
            CheckForOwnerAssociated();

            Worksheet.MergeRange(Position);
        }

        /// <summary>
        ///     Unmerge this range
        /// </summary>
        public void Unmerge()
        {
            CheckForOwnerAssociated();

            Worksheet.UnmergeRange(Position);
        }

        /// <summary>
        ///     Determine whether or not this range contains only one merged cell
        /// </summary>
        public bool IsMergedCell
        {
            get
            {
                CheckForOwnerAssociated();

                var cell = Worksheet.GetCell(StartPos);

                return cell == null ? false : cell.Rowspan == Rows && cell.Colspan == Cols;
            }
        }

#if OUTLINE
		/// <summary>
		/// Group all rows in this range
		/// </summary>
		public void GroupRows()
		{
			CheckForOwnerAssociated();

			this.Worksheet.GroupRows(this.Row, this.Rows);
		}

		/// <summary>
		/// Group all columns in this range
		/// </summary>
		public void GroupColumns()
		{
			CheckForOwnerAssociated();

			this.Worksheet.GroupColumns(this.Col, this.Cols);
		}

		/// <summary>
		/// Ungroup all rows in this range
		/// </summary>
		public void UngroupRows()
		{
			CheckForOwnerAssociated();

			this.Worksheet.UngroupRows(this.Row, this.Rows);
		}

		/// <summary>
		/// Ungroup all columns in this range
		/// </summary>
		public void UngroupColumns()
		{
			CheckForOwnerAssociated();

			this.Worksheet.UngroupColumns(this.Row, this.Rows);
		}
#endif // OUTLINE

        #endregion // Merge and Group

        #region Readonly

        /// <summary>
        ///     Set or get readonly property to all cells inside this range
        /// </summary>
        public bool IsReadonly
        {
            get
            {
                var allReadonly = true;
                foreach (var cell in Cells)
                    if (!cell.IsReadOnly)
                    {
                        allReadonly = false;
                        break;
                    }

                return allReadonly;
            }
            set
            {
                foreach (var cell in Cells) cell.IsReadOnly = value;
            }
        }

        #endregion // Readonly

        private void CheckForOwnerAssociated()
        {
            if (Worksheet == null) throw new ReferenceRangeNotAssociatedException(this);
        }

        #endregion // Control API Routines

        #region Cells Collection

        private Worksheet.CellCollection cellsCollection;

        /// <summary>
        ///     Get the collection of all cell instances in this range
        /// </summary>
        public Worksheet.CellCollection Cells
        {
            get
            {
                if (cellsCollection == null) cellsCollection = new Worksheet.CellCollection(Worksheet, this);

                return cellsCollection;
            }
        }

        #endregion // Cells Collection
    }
}