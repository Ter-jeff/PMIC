﻿#define WPF

using System;
using SpreedSheet.Core;
using SpreedSheet.Interaction;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Events;
#if DEBUG
using System.Diagnostics;
#endif // DEBUG

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        #region Filter

        ///// <summary>
        ///// Create filter on specified column.
        ///// </summary>
        ///// <param name="column">Column code that locates a column to create filter.</param>
        ///// <param name="titleRows">Indicates how many title rows exist at the top of worksheet,
        ///// title rows will not be included in filter range.</param>
        ///// <returns>Instance of column filter.</returns>
        //public AutoColumnFilter CreateColumnFilter(string column, int titleRows = 0,
        //	AutoColumnFilterUI columnFilterUI = AutoColumnFilterUI.DropdownButtonAndPane)
        //{
        //	return CreateColumnFilter(column, column, titleRows, columnFilterUI);
        //}

        /// <summary>
        ///     Create column filter.
        /// </summary>
        /// <param name="startColumn">First column specified by an address to create filter.</param>
        /// <param name="endColumn">Last column specified by an address to the filter.</param>
        /// <param name="titleRows">
        ///     Indicates that how many title rows exist at the top of spreadsheet,
        ///     title rows will not be included in filter apply range.
        /// </param>
        /// <param name="columnFilterUI">
        ///     Indicates whether allow to create graphics user interface (GUI),
        ///     by default the dropdown-button on the column and candidates dropdown-panel will be created.
        ///     Set this argument as NoGUI to create filter without GUI.
        /// </param>
        /// <returns>Instance of column filter.</returns>
        public AutoColumnFilter CreateColumnFilter(string startColumn, string endColumn, int titleRows = 0,
            AutoColumnFilterUI columnFilterUI = AutoColumnFilterUI.DropdownButtonAndPanel)
        {
            var startIndex = RGUtility.GetNumberOfChar(startColumn);
            var endIndex = RGUtility.GetNumberOfChar(endColumn);

            return CreateColumnFilter(startIndex, endIndex, titleRows, columnFilterUI);
        }

        /// <summary>
        ///     Create column filter.
        /// </summary>
        /// <param name="column">Column to create filter.</param>
        /// <param name="titleRows">
        ///     indicates that how many title rows exist at the top of spreadsheet,
        ///     title rows will not be included in filter apply range.
        /// </param>
        /// <param name="columnFilterUI">
        ///     Indicates whether allow to create graphics user interface (GUI),
        ///     by default the dropdown-button on the column and candidates dropdown-panel will be created.
        ///     Set this argument as NoGUI to create filter without GUI.
        /// </param>
        /// <returns>Instance of column filter.</returns>
        public AutoColumnFilter CreateColumnFilter(int column, int titleRows, AutoColumnFilterUI columnFilterUI)
        {
            return CreateColumnFilter(column, column, titleRows, columnFilterUI);
        }

        /// <summary>
        ///     Create column filter.
        /// </summary>
        /// <param name="startColumn">first column specified by a zero-based number of column to create filter</param>
        /// <param name="endColumn">last column specified by a zero-based number of column to create filter</param>
        /// <param name="titleRows">
        ///     indicates that how many title rows exist at the top of spreadsheet,
        ///     title rows will not be included in filter apply range.
        /// </param>
        /// <param name="columnFilterUI">
        ///     Indicates whether or not to show GUI for filter,
        ///     by default the drop-down button displayed on column header and a candidates list popuped up when dropdown-panel
        ///     opened.
        ///     Set this argument as NoGUI to create filter without GUI.
        /// </param>
        /// <returns>Instance of column filter.</returns>
        public AutoColumnFilter CreateColumnFilter(int startColumn, int endColumn, int titleRows = 0,
            AutoColumnFilterUI columnFilterUI = AutoColumnFilterUI.DropdownButtonAndPanel)
        {
            if (startColumn < 0 || startColumn >= ColumnCount)
                throw new ArgumentOutOfRangeException("startColumn",
                    "number of column start to filter out of valid spreadsheet range");

            if (endColumn < startColumn)
                throw new ArgumentOutOfRangeException("endColumn", "end column must be greater than start column");

            if (endColumn >= ColumnCount)
                throw new ArgumentOutOfRangeException("endColumn", "end column out of valid spreadsheet range");

            return CreateColumnFilter(new RangePosition(titleRows, startColumn,
                MaxContentRow - titleRows + 1, endColumn - startColumn + 1), columnFilterUI);
        }

        /// <summary>
        ///     Create automatic column filter and display on specified headers of worksheet.
        /// </summary>
        /// <param name="range">Range to filter data.</param>
        /// <param name="columnFilterUI">
        ///     Indicates whether or not to show GUI for filter,
        ///     by default the drop-down button displayed on column header and a candidates list popuped up when dropdown-panel
        ///     opened.
        ///     Set this argument as NoGUI to create filter without GUI.
        /// </param>
        /// <returns>Instance of column filter.</returns>
        public AutoColumnFilter CreateColumnFilter(RangePosition range,
            AutoColumnFilterUI columnFilterUI = AutoColumnFilterUI.DropdownButtonAndPanel)
        {
            var filter = new AutoColumnFilter(this, FixRange(range));

            filter.Attach(this, columnFilterUI);

            return filter;
        }

        /// <summary>
        ///     Do a filter on specified rows. Determines whether or not to show or hide a row.
        /// </summary>
        /// <param name="startRow">Number of row start to check.</param>
        /// <param name="rows">Number of rows to be checked.</param>
        /// <param name="filter">A callback filter function to check specified row from worksheet.</param>
        public void DoFilter(RangePosition range, Func<int, bool> filter)
        {
            try
            {
                controlAdapter.ChangeCursor(CursorStyle.Busy);

                SetRowsHeight(range.Row, range.Rows, r =>
                {
                    var showRow = filter(r);

                    if (showRow)
                    {
                        var rowhead = RetrieveRowHeader(r);

                        // don't hide row, show the row
                        // if row is hidden, return lastHeight to show the row
                        return rowhead.InnerHeight == 0 ? rowhead.LastHeight : rowhead.InnerHeight;
                    }

                    return 0;
                }, true);
            }
            finally
            {
                if (controlAdapter != null) ControlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
            }

            RowsFiltered?.Invoke(this, null);
        }

        /// <summary>
        ///     Event raised when rows filtered on this worksheet.
        /// </summary>
        public event EventHandler RowsFiltered;

        #endregion // Filter

        #region Sort

        /// <summary>
        ///     Sort data on specified column.
        /// </summary>
        /// <param name="columnAddress">Base column specified by an address to sort data.</param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer">
        ///     Custom cell data comparer, compares two cells and returns an integer.
        ///     Set this value to null to use default built-in comparer.
        /// </param>
        /// <returns>Data changed range</returns>
        public RangePosition SortColumn(string columnAddress, SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            return SortColumn(RGUtility.GetNumberOfChar(columnAddress), order, moveElementFlag, cellDataComparer);
        }

        /// <summary>
        ///     Sort data on specified column.
        /// </summary>
        /// <param name="columnIndex">Zero-based number of column to sort data.</param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer">
        ///     custom cell data comparer, compares two cells and returns an integer.
        ///     Set this value to null to use default built-in comparer.
        /// </param>
        /// <returns>Data changed range</returns>
        public RangePosition SortColumn(int columnIndex, SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            return SortColumn(columnIndex, 0, MaxContentRow, 0, MaxContentCol, order, moveElementFlag,
                cellDataComparer);
        }

        /// <summary>
        ///     Sort data on specified column.
        /// </summary>
        /// <param name="columnIndex">Zero-based number of column to sort data</param>
        /// <param name="titleRows">
        ///     Indicates that how many title rows exist at the top of worksheet,
        ///     Title rows will not be included in sort apply range.
        /// </param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer">
        ///     Custom cell data comparer, compares two cells and returns an integer.
        ///     Set this value to null to use default built-in comparer.
        /// </param>
        /// <returns>Data changed range.</returns>
        public RangePosition SortColumn(int columnIndex, int titleRows, SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            return SortColumn(columnIndex, titleRows, MaxContentRow, 0, MaxContentCol, order, moveElementFlag,
                cellDataComparer);
        }

        /// <summary>
        ///     Sort data on specified column.
        /// </summary>
        /// <param name="columnIndex">Zero-based number of column to sort data.</param>
        /// <param name="startRow">First number of row to allow sort data.</param>
        /// <param name="endRow">Last number of number of row to allow sort data.</param>
        /// <param name="startColumn">First number of column to allow sort data.</param>
        /// <param name="endColumn">Last number of column to allow sort data.</param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer">
        ///     Custom cell data comparer, compares two cells and returns an integer.
        ///     Set this value to null to use default built-in comparer.
        /// </param>
        /// <returns>Data changed range.</returns>
        public RangePosition SortColumn(int columnIndex, int startRow, int endRow, int startColumn, int endColumn,
            SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            return SortColumn(columnIndex, new RangePosition(startRow, startColumn, MaxContentRow - startRow + 1,
                endColumn - startColumn + 1), order, moveElementFlag, cellDataComparer);
        }

        /// <summary>
        ///     Sort data inside specified range.
        /// </summary>
        /// <param name="columnIndex">Data will be sorted based on this column.</param>
        /// <param name="applyRange">Affect range.</param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer"></param>
        /// <returns></returns>
        public RangePosition SortColumn(int columnIndex, string applyRange, SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            if (RangePosition.IsValidAddress(applyRange))
                return SortColumn(columnIndex, new RangePosition(applyRange), order, moveElementFlag, cellDataComparer);

            RangePosition range;
            if (TryGetNamedRangePosition(applyRange, out range))
                return SortColumn(columnIndex, range, order, moveElementFlag, cellDataComparer);
            throw new InvalidAddressException(applyRange);
        }

        /// <summary>
        ///     Sort data on specified column.
        /// </summary>
        /// <param name="columnIndex">Zero-based number of column to sort data.</param>
        /// <param name="applyRange">Data only be changed in this range during sort.</param>
        /// <param name="order">Order of data sort.</param>
        /// <param name="cellDataComparer">
        ///     Custom cell data comparer, compares two cells and returns an integer.
        ///     Set this value to null to use default built-in comparer.
        /// </param>
        /// <returns>Data changed range.</returns>
        public RangePosition SortColumn(int columnIndex, RangePosition applyRange,
            SortOrder order = SortOrder.Ascending,
            CellElementFlag moveElementFlag = CellElementFlag.Data,
            Func<int, int, object, object, int> cellDataComparer = null)
        {
            var range = FixRange(applyRange);

            var affectRange = RangePosition.Empty;

            if (cellDataComparer != null) this.cellDataComparer = cellDataComparer;

#if DEBUG
            var sw = Stopwatch.StartNew();
#endif // DEBUG

            int[] sortedOrder = null;

            // stop fire events
            SuspendDataChangedEvents();

            try
            {
                controlAdapter.ChangeCursor(CursorStyle.Busy);

                if (!CheckQuickSortRange(columnIndex, range.Row, range.EndRow, range.Col, range.EndCol, order,
                        ref affectRange))
                    throw new InvalidOperationException(
                        "Cannot change a part of range, all cells should be having same colspan on column.");

                QuickSortColumn(columnIndex, range.Row, range.EndRow, range.Col, range.EndCol, order, ref affectRange,
                    cellDataComparer == null ? (Func<int, int, object, int>)CompareCell : UserCellDataComparerAdapter,
                    sortedOrder);

#if DEBUG
                sw.Stop();

                if (sw.ElapsedMilliseconds > 10)
                    Debug.WriteLine("sort column by {0} on [{1}-{2}]: {3} ms", columnIndex, range.Col, range.EndCol,
                        sw.ElapsedMilliseconds);
#endif // DEBUG
            }
            finally
            {
                // resume to fire events
                ResumeDataChangedEvents();
            }

            RequestInvalidate();

            controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);

            if (!affectRange.IsEmpty)
            {
                for (var c = affectRange.Col; c <= affectRange.EndCol; c++)
                {
                    var header = cols[c];

                    if (header.Body != null) header.Body.OnDataChange(affectRange.Row, affectRange.EndRow);
                }

                RaiseRangeDataChangedEvent(affectRange);

                RowsSorted?.Invoke(this, new RangeEventArgs(affectRange));
            }

            return affectRange;
        }

        private bool CheckQuickSortRange(int columnIndex, int row, int endRow, int col, int endCol,
            SortOrder order, ref RangePosition affectRange)
        {
            for (var c = col; c <= endCol; c++)
            {
                var cell1 = cells[row, c];

                for (var r = row + 1; r <= endRow; r++)
                {
                    var cell2 = cells[r, c];

                    if (cell1 != null && cell2 != null
                                      && ((cell1.IsValidCell && !cell2.IsValidCell)
                                          || (!cell1.IsValidCell && cell2.IsValidCell)))
                        return false;
                }
            }

            return true;
        }

        private void QuickSortRelocateRow(int row, int[] rowIndexes, int startColumn, int endColumn, int excludeColumn)
        {
            var targetIndex = rowIndexes[row];
            //int targetTargetIndex = rowIndexes[targetIndex];

            QuickSortSwapRow(row, targetIndex, rowIndexes, startColumn, endColumn, excludeColumn);

            var index = rowIndexes[row];
            rowIndexes[row] = rowIndexes[targetIndex];
            rowIndexes[targetIndex] = index;
        }

        private void QuickSortSwapRow(int row1, int row2, int[] rowIndexes, int startColumn, int endColumn,
            int excludeColumn)
        {
            for (var col = startColumn; col <= endColumn; col++)
            {
                if (col == excludeColumn) continue;

                var cell1 = cells[row1, col];
                var cell2 = cells[row2, col];

                if ((cell1.IsValidCell && !cell2.IsValidCell)
                    || (!cell1.IsValidCell && cell2.IsValidCell))
                    throw new InvalidOperationException(
                        "Cannot change a part of range, all cells should be having same colspans during sort.");

                if (cell1.InnerData != null || cell2.InnerData != null)
                {
                    var v = cell1.InnerData;
                    SetSingleCellData(cell1, cell2.InnerData);
                    SetSingleCellData(cell2, v);
                }
            }
        }

        private int UserCellDataComparerAdapter(int row, int col, object @base)
        {
            if (@base == null) return 0;

            var data = GetCellData(row, col) as IComparable;
            if (data == null) return 0;

            return cellDataComparer(row, col, data, @base);
        }

        private Func<int, int, object, object, int> cellDataComparer;

        private int CompareCell(int row, int col, object @base)
        {
            if (@base == null) return 0;

            var data = GetCellData(row, col) as IComparable;
            if (data == null) return 0;

            if (data.GetType() == @base.GetType())
                return data.CompareTo(@base);
            if (@base is string)
                return Convert.ToString(data).CompareTo(@base);
            if (data is string)
                return data.CompareTo(Convert.ToString(@base));
            try
            {
                return ((double)Convert.ChangeType(data, typeof(double))).CompareTo(
                    Convert.ChangeType(@base, typeof(double)));
            }
            catch
            {
                return Convert.ToString(data).CompareTo(Convert.ToString(@base));
            }
        }

        private void QuickSortColumn(int columnIndex, int start, int end, int startColumn, int endColumn,
            SortOrder order,
            ref RangePosition affectRange, Func<int, int, object, int> cellComparer, int[] sortedOrder)
        {
            while (start < end)
            {
                object @base = GetCellData((start + end) / 2, columnIndex) as IComparable;

                int top = start, bottom = end;

                while (true)
                {
                    if (order == SortOrder.Ascending)
                    {
                        while (cellComparer(top, columnIndex, @base) < 0) top++;
                        while (cellComparer(bottom, columnIndex, @base) > 0) bottom--;
                    }
                    else
                    {
                        while (cellComparer(top, columnIndex, @base) > 0) top++;
                        while (cellComparer(bottom, columnIndex, @base) < 0) bottom--;
                    }

                    if (top >= bottom) break;

                    for (var col = startColumn; col <= endColumn; col++)
                    {
                        var v = GetCellData(top, col);
                        SetCellData(top, col, GetCellData(bottom, col));
                        SetCellData(bottom, col, v);
                    }

                    if (affectRange.IsEmpty)
                    {
                        affectRange.Row = start;
                        affectRange.Col = startColumn;
                        affectRange.EndRow = end;
                        affectRange.EndCol = endColumn;
                    }
                    else
                    {
                        if (affectRange.Row > start) affectRange.Row = start;
                        if (affectRange.EndRow < end) affectRange.EndRow = end;
                    }

                    top++;
                    bottom--;
                }

                QuickSortColumn(columnIndex, start, top - 1, startColumn, endColumn, order,
                    ref affectRange, cellComparer, sortedOrder);

                start = bottom + 1;
            }
        }

        /// <summary>
        ///     Event raised when rows sorted on this worksheet.
        /// </summary>
        public event EventHandler<RangeEventArgs> RowsSorted;

        #endregion // Sort
    }

    /// <summary>
    ///     Sort order.
    /// </summary>
    public enum SortOrder
    {
        /// <summary>
        ///     Ascending
        /// </summary>
        Ascending,

        /// <summary>
        ///     Descending
        /// </summary>
        Descending
    }
}