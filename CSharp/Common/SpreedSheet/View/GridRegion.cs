using SpreedSheet.Core;

namespace SpreedSheet.View
{
    internal struct GridRegion
    {
        internal int StartRow;
        internal int EndRow;
        internal int StartCol;
        internal int EndCol;

        internal static readonly GridRegion Empty = new GridRegion
        {
            StartRow = 0,
            StartCol = 0,
            EndRow = 0,
            EndCol = 0
        };

        public GridRegion(int startRow, int startCol, int endRow, int endCol)
        {
            StartRow = startRow;
            StartCol = startCol;
            EndRow = endRow;
            EndCol = endCol;
        }

        public bool Contains(CellPosition pos)
        {
            return Contains(pos.Row, pos.Col);
        }

        public bool Contains(int row, int col)
        {
            return StartRow <= row && EndRow >= row && StartCol <= col && EndCol >= col;
        }

        public bool Contains(RangePosition range)
        {
            return range.Row >= StartRow && range.Col >= StartCol
                                         && range.EndRow <= EndRow && range.EndCol <= EndCol;
        }

        public bool Intersect(RangePosition range)
        {
            return (range.Row < StartRow && range.EndRow > StartRow)
                   || (range.Row < EndRow && range.EndRow > EndRow)
                   || (range.Col < StartCol && range.EndCol > StartCol)
                   || (range.Col < EndCol && range.EndCol > EndCol);
        }

        public bool IsOverlay(RangePosition range)
        {
            return Contains(range) || Intersect(range);
        }

        public override bool Equals(object obj)
        {
            if (obj as GridRegion? == null) return false;

            var gr2 = (GridRegion)obj;
            return StartRow == gr2.StartRow && StartCol == gr2.StartCol
                                            && EndRow == gr2.EndRow && EndCol == gr2.EndCol;
        }

        public override int GetHashCode()
        {
            return StartRow ^ StartCol ^ EndRow ^ EndCol;
        }

        public static bool operator ==(GridRegion gr1, GridRegion gr2)
        {
            return gr1.Equals(gr2);
        }

        public static bool operator !=(GridRegion gr1, GridRegion gr2)
        {
            return !gr1.Equals(gr2);
        }

        public bool IsEmpty
        {
            get { return Equals(Empty); }
        }

        public int Rows
        {
            get { return EndRow - StartRow + 1; }
            set { EndRow = StartRow + value - 1; }
        }

        public int Cols
        {
            get { return EndCol - StartCol + 1; }
            set { EndCol = StartCol + value - 1; }
        }

        public override string ToString()
        {
            return string.Format("VisibleRegion[{0},{1}-{2},{3}]", StartRow, StartCol, EndRow, EndCol);
        }

        /// <summary>
        ///     Convert into range struct
        /// </summary>
        /// <returns></returns>
        public RangePosition ToRange()
        {
            return new RangePosition(StartRow, StartCol, EndRow - StartRow + 1, EndCol - StartCol + 1);
        }
    }
}