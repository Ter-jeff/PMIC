#define WPF


#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using System;
using SpreedSheet.Core;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Graphics;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WPF

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        internal Rectangle GetRangeBounds(int row, int col, int rows, int cols)
        {
            return GetRangePhysicsBounds(new RangePosition(row, col, rows, cols));
        }

        internal Rectangle GetRangeBounds(CellPosition startPos, CellPosition endPos)
        {
            return GetRangePhysicsBounds(new RangePosition(startPos, endPos));
        }

        /// <summary>
        ///     Get physics rectangle bounds from specified range position.
        ///     Be careful that this is different from the rectangle bounds displayed on screen,
        ///     the actual bound positions displayed on screen are transformed and scaled
        ///     in order to scroll, zoom and freeze into different viewports.
        /// </summary>
        /// <param name="range">The range position to get bounds</param>
        /// <returns>Rectangle bounds from specified range position</returns>
        public Rectangle GetRangePhysicsBounds(RangePosition range)
        {
            var fixedRange = FixRange(range);

            var rowHead = rows[fixedRange.Row];
            var colHead = cols[fixedRange.Col];
            var toRowHead = rows[fixedRange.EndRow];
            var toColHead = cols[fixedRange.EndCol];

            var width = toColHead.Right - colHead.Left;
            var height = toRowHead.Bottom - rowHead.Top;

            return new Rectangle(colHead.Left, rowHead.Top, width + 1, height + 1);
        }

        /// <summary>
        ///     Get physics position from specified cell position.
        ///     Be careful that this is different from the rectangle bounds displayed on screen,
        ///     the actual bound positions displayed on the screen are transformed and scaled
        ///     in order to scroll, zoom and freeze into different viewports.
        /// </summary>
        /// <param name="row">Zero-based index of row</param>
        /// <param name="col">Zero-based index of column</param>
        /// <returns>Point position of specified cell position in pixel.</returns>
        public Point GetCellPhysicsPosition(int row, int col)
        {
            if (row < 0 || row >= rows.Count
                        || col < 0 || col >= cols.Count)
                throw new ArgumentException("row or col invalid");

            var rowHeader = rows[row].Top;
            var colHeader = cols[col].Left;

            return new Point(colHeader, rowHeader);
        }

        internal Rectangle GetScaledRangeBounds(RangePosition range)
        {
            var rowHead = rows[range.Row];
            var colHead = cols[range.Col];
            var toRowHead = rows[range.EndRow];
            var toColHead = cols[range.EndCol];

            var width = (toColHead.Right - colHead.Left) * renderScaleFactor;
            var height = (toRowHead.Bottom - rowHead.Top) * renderScaleFactor;

            return new Rectangle(colHead.Left * renderScaleFactor, rowHead.Top * renderScaleFactor, width, height);
        }


        internal Rectangle GetCellBounds(CellPosition pos)
        {
            return GetCellBounds(pos.Row, pos.Col);
        }

        internal Rectangle GetCellBounds(int row, int col)
        {
            if (cells[row, col] == null) return GetCellRectFromHeader(row, col);

            if (cells[row, col].MergeStartPos != CellPosition.Empty)
            {
                var cell = GetCell(cells[row, col].MergeStartPos);
                return cell?.Bounds ?? GetCellRectFromHeader(row, col);
            }

            return cells[row, col].Bounds;
        }

        private Rectangle GetCellRectFromHeader(int row, int col)
        {
            return new Rectangle(cols[col].Left, rows[row].Top, cols[col].InnerWidth + 1, rows[row].InnerHeight + 1);
        }

        #region Header

        internal int FindColIndexMiddle(double x)
        {
            return ArrayHelper.QuickFind(0, cols.Count, i =>
            {
                var col = cols[i];

                if (x > col.Left + col.InnerWidth / 2)
                    return 1;

                if (i > 0)
                {
                    var prevCol = cols[i - 1];

                    if (x < prevCol.Left + prevCol.InnerWidth / 2) return -1;
                }

                return 0;
            });
        }

        internal int FindRowIndexMiddle(double x)
        {
            return ArrayHelper.QuickFind(0, rows.Count, i =>
            {
                var row = rows[i];

                if (x > row.Top + row.InnerHeight / 2)
                    return 1;

                if (i > 0)
                {
                    var prevCol = rows[i - 1];

                    if (x < prevCol.Top + prevCol.InnerHeight / 2) return -1;
                }

                return 0;
            });
        }

        // TODO: need performance improvement
        internal bool FindColumnByPosition(double x, out int col)
        {
            var v = -1;
            var inline = true;

            var scaleThumb = 2 / renderScaleFactor;

            for (var i = 0; i < cols.Count; i++)
                if (x <= cols[i].Right - scaleThumb)
                {
                    inline = false;
                    v = i;
                    break;
                }
                else if (x <= cols[i].Right + scaleThumb)
                {
                    v = i;
                    break;
                }

            col = v;
            return inline;
        }

        // TODO: need performance improvement
        internal bool FindRowByPosition(double y, out int row)
        {
            var v = -1;
            var inline = true;

            var scaleThumb = 2 / renderScaleFactor;

            for (var i = 0; i < rows.Count; i++)
                if (y <= rows[i].Bottom - scaleThumb)
                {
                    inline = false;
                    v = i;
                    break;
                }
                else if (y <= rows[i].Bottom + scaleThumb)
                {
                    v = i;
                    break;
                }

            row = v;
            return inline;
        }

        #endregion // Header
    }
}