#define WPF

using System;
using System.ComponentModel;
using SpreedSheet.Core;
using SpreedSheet.View;
using SpreedSheet.View.Controllers;
using unvell.Common;
using unvell.Common.Win32Lib;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
#if DEBUG
using System.Diagnostics;
#endif // DEBUG

#if EX_SCRIPT
using unvell.ReoScript;
#endif // EX_SCRIPT

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        #region Position

        internal CellPosition selStart = new CellPosition(0, 0);
        internal CellPosition selEnd = new CellPosition(0, 0);

        #region Focus & Hover

        internal CellPosition focusPos = new CellPosition(0, 0);

        /// <summary>
        ///     The column focus pos goes when enter key pressed.
        /// </summary>
        private int focusReturnColumn;

        /// <summary>
        ///     Get or set current focused cell position.
        /// </summary>
        public CellPosition FocusPos
        {
            get { return focusPos; }
            set
            {
                // different with current focus pos
                if (focusPos != value)
                {
                    var newFocusPos = FixPos(value);

                    // not empty position
                    if (!newFocusPos.IsEmpty)
                    {
                        var focusCell = cells[newFocusPos.Row, newFocusPos.Col];

                        if (focusCell != null)
                            // new focus cell may be an invalid cell, need check it
                            if (!focusCell.IsValidCell)
                                // if inside any merged cell, find the merge-start-cell
                                newFocusPos = GetMergedCellOfRange(focusCell).InternalPos;
                    }

                    // compare to current focus position again
                    if (focusPos != newFocusPos)
                    {
                        // if current focus position is not empty
                        if (!focusPos.IsEmpty)
                        {
                            // get the cell, and invoke OnLostFocus if cell's has body
                            var focusCell = cells[focusPos.Row, focusPos.Col];

                            if (focusCell != null && focusCell.body != null) focusCell.body.OnLostFocus();
                        }

                        focusPos = newFocusPos;

                        // invoke OnGotFocus on new focus position
                        if (!focusPos.IsEmpty)
                        {
                            var focusCell = cells[focusPos.Row, focusPos.Col];

                            if (focusCell != null && focusCell.body != null && focusCell.IsValidCell)
                                focusCell.body.OnGotFocus();

                            if (!selectionRange.Contains(focusPos)) SelectRange(focusPos.Row, FocusPos.Col, 1, 1);
                        }

                        RequestInvalidate();

                        FocusPosChanged?.Invoke(this, new CellPosEventArgs(focusPos));
                    }
                }
            }
        }

        /// <summary>
        ///     Raise when focus cell is changed
        /// </summary>
        public event EventHandler<CellPosEventArgs> FocusPosChanged;

        private FocusPosStyle focusPosStyle = FocusPosStyle.Default;

        /// <summary>
        ///     Get or set focus position display style
        /// </summary>
        public FocusPosStyle FocusPosStyle
        {
            get { return focusPosStyle; }
            set
            {
                if (focusPosStyle != value)
                {
                    RequestInvalidate();

                    focusPosStyle = value;

                    FocusPosStyleChanged?.Invoke(this, null);
                }
            }
        }

        /// <summary>
        ///     Focus position style changed.
        /// </summary>
        public event EventHandler<EventArgs> FocusPosStyleChanged;

        internal CellPosition hoverPos;

        /// <summary>
        ///     Cell when mouse moving and hover on
        /// </summary>
        public CellPosition HoverPos
        {
            get { return hoverPos; }

            internal set
            {
                if (hoverPos != value)
                {
                    // raise cell mouse enter
                    if (!hoverPos.IsEmpty)
                    {
                        CellMouseEventArgs evtArg = null;

                        if (CellMouseLeave != null)
                        {
                            evtArg = new CellMouseEventArgs(this, hoverPos);
                            CellMouseLeave(this, evtArg);
                        }

                        var cell = cells[hoverPos.Row, hoverPos.Col];

                        if (cell != null)
                        {
                            if (!cell.IsValidCell) cell = GetMergedCellOfRange(cell);

                            if (cell.body != null)
                            {
                                if (evtArg == null) evtArg = new CellMouseEventArgs(this, cell);

                                var processed = cell.body.OnMouseLeave(evtArg);
                                if (processed) RequestInvalidate();
                            }
                        }
                    }

                    hoverPos = value;

                    // raise cell mouse leave
                    if (!hoverPos.IsEmpty)
                    {
                        CellMouseEventArgs evtArg = null;

                        if (CellMouseEnter != null)
                        {
                            evtArg = new CellMouseEventArgs(this, hoverPos);
                            CellMouseEnter(this, evtArg);
                        }

                        var cell = cells[hoverPos.Row, hoverPos.Col];

                        if (cell != null)
                        {
                            if (!cell.IsValidCell) cell = GetMergedCellOfRange(cell);

                            if (cell.body != null)
                            {
                                if (evtArg == null)
                                {
                                    evtArg = new CellMouseEventArgs(this, cell);
                                    evtArg.Cell = cell;
                                }

                                var processed = cell.body.OnMouseEnter(evtArg);
                                if (processed) RequestInvalidate();
                            }
                        }
                    }

                    HoverPosChanged?.Invoke(this, new CellPosEventArgs(hoverPos));
                }
            }
        }

        /// <summary>
        ///     Raise when hover cell is changed
        /// </summary>
        public event EventHandler<CellPosEventArgs> HoverPosChanged;

        #endregion // Focus & Hover

        internal RangePosition selectionRange = new RangePosition(0, 0, 1, 1);

        /// <summary>
        ///     Current selection range of entire grid. If SelectionMode is None, the value of this property will be Empty.
        /// </summary>
        public RangePosition SelectionRange
        {
            get { return selectionRange; }
            set { SelectRange(value); }
        }

        #endregion // Position

        #region Mode & Style

        internal WorksheetSelectionMode selectionMode = WorksheetSelectionMode.Range;

        /// <summary>
        ///     Get or set selection mode for worksheet.
        /// </summary>
        [DefaultValue(WorksheetSelectionMode.Range)]
        public WorksheetSelectionMode SelectionMode
        {
            get { return selectionMode; }
            set
            {
                if (selectionMode != value)
                {
                    if (IsEditing) EndEdit(EndEditReason.NormalFinish);

                    var oldSelectionMode = selectionMode;

                    selectionMode = value;

                    switch (oldSelectionMode)
                    {
                        case WorksheetSelectionMode.None:
                            switch (value)
                            {
                                case WorksheetSelectionMode.Cell:
                                case WorksheetSelectionMode.Range:

                                    #region None -> Cell/Range

                                    SelectRange(new RangePosition(0, 0, 1, 1));

                                    #endregion // None -> Cell/Range

                                    break;
                            }

                            break;

                        default:
                            switch (value)
                            {
                                case WorksheetSelectionMode.None:

                                    #region Any -> None

                                    selectionRange = RangePosition.Empty;
                                    focusPos = CellPosition.Empty;
                                    RequestInvalidate();

                                    #endregion // Any -> None

                                    break;

                                case WorksheetSelectionMode.Cell:

                                    #region Any -> Cell

                                    SelectRange(selStart.Row, selStart.Col, 1, 1);

                                    #endregion // Any -> Cell

                                    break;

                                case WorksheetSelectionMode.Range:
                                    SelectionRange = FixRangeSelection(selectionRange);
                                    break;
                            }

                            break;
                    }

                    switch (selectionMode)
                    {
                        case WorksheetSelectionMode.Row:
                        case WorksheetSelectionMode.SingleRow:

                            #region Any -> Row

                            SelectRange(selectionRange.Row, 0, selectionRange.Rows, -1);

                            #endregion // None -> Row

                            break;

                        case WorksheetSelectionMode.Column:
                        case WorksheetSelectionMode.SingleColumn:

                            #region Any -> Column

                            SelectRange(0, selectionRange.Col, -1, selectionRange.Cols);

                            #endregion // None -> Column

                            break;
                    }

                    if (SelectionModeChanged != null) SelectionModeChanged(this, null);
                }
            }
        }

        private WorksheetSelectionStyle selectionStyle = WorksheetSelectionStyle.Default;

        /// <summary>
        ///     Get or set the selection style for worksheet.
        /// </summary>
        [DefaultValue(WorksheetSelectionStyle.Default)]
        public WorksheetSelectionStyle SelectionStyle
        {
            get { return selectionStyle; }
            set
            {
                if (selectionStyle != value)
                {
                    selectionStyle = value;
                    RequestInvalidate();

                    SelectionStyleChanged?.Invoke(this, null);
                }
            }
        }

        private SelectionForwardDirection selectionForwardDirection;

        /// <summary>
        ///     Get or set focus cell moving direction.
        /// </summary>
        [DefaultValue(SelectionForwardDirection.Right)]
        public SelectionForwardDirection SelectionForwardDirection
        {
            get { return selectionForwardDirection; }
            set
            {
                if (selectionForwardDirection != value)
                {
                    selectionForwardDirection = value;

                    SelectionForwardDirectionChanged?.Invoke(this, null);
                }
            }
        }

        #endregion // Mode & Style

        #region Mouse Select

        internal void SelectRangeStartByMouse(Point location)
        {
            if (ViewportController == null || ViewportController.View == null) return;

#if WINFORM || WPF
            if (!Toolkit.IsKeyDown(Win32.VKey.VK_SHIFT))
            {
#endif // WINFORM || WPF
                var viewport = ViewportController.View.GetViewByPoint(location) as IRangeSelectableView;

                if (viewport == null) viewport = ViewportController.FocusView as IRangeSelectableView;

                if (viewport != null)
                {
                    var vp = viewport.PointToView(location);

                    var pos = CellsViewport.GetPosByPoint(viewport, vp);
                    selEnd = selStart = pos;
                }
#if WINFORM || WPF
            }
#endif // WINFORM || WPF

            SelectRangeEndByMouse(location);
        }

        internal void SelectRangeEndByMouse(Point location)
        {
            if (ViewportController == null || ViewportController.View == null) return;

            var viewport = ViewportController.View.GetViewByPoint(location) as IRangeSelectableView;

            if (viewport == null) viewport = ViewportController.FocusView as IRangeSelectableView;

            if (viewport != null)
            {
                var vp = viewport.PointToView(location);

                var startpos = selStart;
                var endpos = selEnd;

                #region Each Operation Status

                switch (operationStatus)
                {
                    case OperationStatus.FullColumnSelect:
                    {
                        var col = -1;

                        FindColumnByPosition(vp.X, out col);

                        if (col > -1)
                        {
                            startpos = new CellPosition(0, startpos.Col);
                            endpos = new CellPosition(rows.Count, col);
                        }
                    }
                        break;

                    case OperationStatus.FullRowSelect:
                    {
                        var row = -1;

                        FindRowByPosition(vp.Y, out row);

                        if (row > -1)
                        {
                            startpos = new CellPosition(startpos.Row, 0);
                            endpos = new CellPosition(row, cols.Count);
                        }
                    }
                        break;

                    default:
                        endpos = CellsViewport.GetPosByPoint(viewport, vp);
                        break;
                }

                #endregion // Each Operation Status

                ApplyRangeSelection(startpos, endpos);
            }
        }

        #endregion // Mouse Select

        #region Select API

        /// <summary>
        ///     Select specified range.
        /// </summary>
        /// <param name="range">Specified range to be selected</param>
        private RangePosition FixRangeSelection(RangePosition range)
        {
            if (range.IsEmpty) return RangePosition.Empty;

#if DEBUG
            var stop = Stopwatch.StartNew();
#endif

            var fixedRange = FixRange(range);

            var minr = fixedRange.Row;
            var minc = fixedRange.Col;
            var maxr = fixedRange.EndRow;
            var maxc = fixedRange.EndCol;

            switch (selectionMode)
            {
                case WorksheetSelectionMode.Cell:
                    maxr = minr = range.Row;
                    maxc = minc = range.Col;
                    break;

                case WorksheetSelectionMode.Row:
                    minc = 0;
                    maxc = this.cols.Count - 1;
                    break;

                case WorksheetSelectionMode.Column:
                    minr = 0;
                    maxr = this.rows.Count - 1;
                    break;
            }

            if ((selectionMode == WorksheetSelectionMode.Cell
                 || selectionMode == WorksheetSelectionMode.Range)
                && ((fixedRange.Cols < this.cols.Count
                     && fixedRange.Rows < this.rows.Count)
                    || this.cols.Count == 1 || this.rows.Count == 1)
               )
            {
                #region Check and select the whole merged region

                //#if DEBUG
                //				if (!Toolkit.IsKeyDown(unvell.Common.Win32Lib.Win32.VKey.VK_CONTROL))
                //				{
                //#endif
                //
                // if there are any entire rows or columns selected (full == -1)
                // the selection bounds of merged range will not be checked.
                // any changes to the selection will also not be appiled to the range.
                //
                var checkedRange = CheckMergedRange(new RangePosition(minr, minc, maxr - minr + 1, maxc - minc + 1));

                minr = checkedRange.Row;
                minc = checkedRange.Col;
                maxr = checkedRange.EndRow;
                maxc = checkedRange.EndCol;

                //#if DEBUG
                //				}
                //#endif

                #endregion
            }

            var rows = maxr - minr + 1;
            var cols = maxc - minc + 1;

#if DEBUG
            stop.Stop();
            if (stop.ElapsedMilliseconds > 25)
                Debug.WriteLine("select range takes " + stop.ElapsedMilliseconds + " ms.");
#endif

            return new RangePosition(minr, minc, rows, cols);
        }

        private void MoveRangeSelection(CellPosition start, CellPosition end, bool appendSelect,
            bool scrollToSelectionEnd = true)
        {
            if (!appendSelect) start = end;

            ApplyRangeSelection(start, end, scrollToSelectionEnd);
        }

        private void ApplyRangeSelection(CellPosition start, CellPosition end, bool scrollToSelectionEnd = true)
        {
            var processed = false;

            switch (operationStatus)
            {
                case OperationStatus.HighlightRangeCreate:
                {
                    var fixedRange = FixRangeSelection(new RangePosition(start, end));

                    if (focusHighlightRange == null)
                    {
                        // no focus highlight range, create one
                        var refRange = AddHighlightRange(fixedRange);
                        FocusHighlightRange = refRange;
                    }
                    else if (focusHighlightRange.Position != fixedRange)
                    {
                        // update size for current focus highlight range
                        focusHighlightRange.Position = fixedRange;
                        RequestInvalidate();
                    }
                }

                    processed = true;
                    break;

                default:
                    ChangeSelectionRange(start, end);

                    processed = true;
                    break;
            }

            if (processed)
                if (HasSettings(WorksheetSettings.Behavior_ScrollToFocusCell)
                    //commented out before the case of entire row or column
                    //is checked inside NormalViewportController.ScrollToRange method
                    //issue #179
                    //&& (this.selectionRange.Rows != this.rows.Count
                    //&& this.selectionRange.Cols != this.cols.Count)
                    && scrollToSelectionEnd
                   )
                    // skip to scroll if entire worksheet is selected
                    if (!(start.Row == 0 && start.Col == 0
                                         && selEnd.Row == rows.Count - 1 && selEnd.Col == cols.Count - 1))
                        ScrollToCell(selEnd);
        }

        private void ChangeSelectionRange(CellPosition start, CellPosition end)
        {
            var range = FixRangeSelection(new RangePosition(start, end));

            // compare to current selection, only do this when selection was really changed.
            if (selectionRange != range)
            {
                if (BeforeSelectionRangeChange != null)
                {
                    var arg = new BeforeSelectionChangeEventArgs(start, end);
                    BeforeSelectionRangeChange(this, arg);

                    if (arg.IsCancelled) return;

                    if (start != arg.SelectionStart || end != arg.SelectionEnd)
                    {
                        start = arg.SelectionStart;
                        end = arg.SelectionEnd;

                        range = FixRangeSelection(new RangePosition(start, end));
                    }
                }

                selectionRange = range;

                selStart = start;
                selEnd = end;

                //if (!range.Contains(selStart)) selStart = range.StartPos;
                //if (!range.Contains(selEnd)) selEnd = range.EndPos;

                // focus pos validations:
                //   1. focus pos must be inside selection range
                //   2. focus pos cannot stop at invalid cell (any part of merged cell)
                if (this.focusPos.IsEmpty
                    || !range.Contains(this.focusPos)
                    || !IsValidCell(this.focusPos))
                {
                    var focusPos = selStart;

                    // find first valid cell as focus pos
                    for (var r = range.Row; r <= range.EndRow; r++)
                    for (var c = range.Col; c <= range.EndCol; c++)
                    {
                        var cell = cells[r, c];
                        if (cell != null && (cell.Colspan <= 0 || cell.Rowspan <= 0)) continue;

                        focusPos.Row = r;
                        focusPos.Col = c;
                        goto quit_loop;
                    }

                    quit_loop:

                    if (focusPos.Col < cols.Count
                        && focusPos.Row < rows.Count)
                        FocusPos = focusPos;
                }

                // update focus return column
                focusReturnColumn = end.Col;

                if (operationStatus == OperationStatus.RangeSelect)
                {
                    SelectionRangeChanging?.Invoke(this, new RangeEventArgs(selectionRange));

#if EX_SCRIPT
					// comment out this if you get performance problem when using script extension
					RaiseScriptEvent("onselectionchanging");
#endif
                }
                else
                {
                    SelectionRangeChanged?.Invoke(this, new RangeEventArgs(selectionRange));

#if EX_SCRIPT
					RaiseScriptEvent("onselectionchange");
#endif
                }

                RequestInvalidate();
            }
        }

        /// <summary>
        ///     Select speicifed range on spreadsheet
        /// </summary>
        /// <param name="address">address or name of specified range to be selected</param>
        public void SelectRange(string address)
        {
            // range address
            if (RangePosition.IsValidAddress(address))
            {
                SelectRange(new RangePosition(address));
            }
            // named range
            else if (RGUtility.IsValidName(address))
            {
                NamedRange refRange;
                if (registeredNamedRanges.TryGetValue(address, out refRange)) SelectRange(refRange);
            }
        }

        /// <summary>
        ///     Select speicifed range on spreadsheet
        /// </summary>
        /// <param name="pos1">Start position of specified range</param>
        /// <param name="pos2">End position of specified range</param>
        public void SelectRange(CellPosition pos1, CellPosition pos2)
        {
            SelectRange(new RangePosition(pos1, pos2));
        }

        /// <summary>
        ///     Select specified range
        /// </summary>
        /// <param name="row">number of row</param>
        /// <param name="col">number of col</param>
        /// <param name="rows">number of rows to be selected</param>
        /// <param name="cols">number of columns to be selected</param>
        public void SelectRange(int row, int col, int rows, int cols)
        {
            SelectRange(new RangePosition(row, col, rows, cols));
        }

        /// <summary>
        ///     Select speicifed range on spreadsheet
        /// </summary>
        /// <param name="range">range to be selected</param>
        public void SelectRange(RangePosition range)
        {
            if (range.IsEmpty || selectionMode == WorksheetSelectionMode.None) return;

            range = FixRange(range);

            // submit to select a range 
            ApplyRangeSelection(range.StartPos, range.EndPos, false);
        }

        /// <summary>
        ///     Select entire sheet
        /// </summary>
        public void SelectAll()
        {
            if (IsEditing)
                controlAdapter.EditControlSelectAll();
            else
                SelectRange(new RangePosition(0, 0, RowCount, ColumnCount));
        }

        /// <summary>
        ///     Select entire rows of columns form specified column
        /// </summary>
        /// <param name="col">number of column start to be selected</param>
        /// <param name="columns">numbers of column to be selected</param>
        public void SelectColumns(int col, int columns)
        {
            SelectRange(new RangePosition(0, col, rows.Count, columns));
        }

        /// <summary>
        ///     Select entire column of rows from specified row
        /// </summary>
        /// <param name="row">number of row start to be selected</param>
        /// <param name="rows">numbers of row to be selected</param>
        public void SelectRows(int row, int rows)
        {
            SelectRange(new RangePosition(row, 0, rows, cols.Count));
        }

        #endregion // Select API

        #region Keyboard Move

        private void OnTabKeyPressed(bool shiftKeyDown)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var backupReturnCol = focusReturnColumn;

            if (!shiftKeyDown)
            {
                var endCol = selectionRange.Cols > 1
                    ? selectionRange.EndCol
                    : cols.Count - 1;

                if (focusPos.Col < endCol)
                {
                    MoveFocusRight();
                }
                else
                {
                    var endRow = selectionRange.Rows > 1
                        ? selectionRange.EndRow
                        : rows.Count - 1;

                    if (focusPos.Row < endRow)
                    {
                        var startCol = selectionRange.Cols > 1 ? selectionRange.Col : 0;

                        focusPos.Col = startCol;

                        MoveFocusDown();
                    }
                }
            }
            else
            {
                var startCol = selectionRange.Cols > 1 ? selectionRange.Col : 0;

                if (selEnd.Col > startCol)
                {
                    MoveSelectionLeft();
                }
                else
                {
                    var startRow = selectionRange.Rows > 1 ? selectionRange.Row : 0;

                    if (selEnd.Row > startRow)
                    {
                        var endCol = selectionRange.Cols > 1
                            ? selectionRange.EndCol
                            : cols.Count - 1;

                        focusPos.Col = endCol;

                        MoveSelectionUp();
                    }
                }
            }

            focusReturnColumn = backupReturnCol;
        }

        private void OnEnterKeyPressed(bool shiftKeyDown)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            if (!shiftKeyDown)
                MoveSelectionForward();
            else
                MoveSelectionBackward();
        }

        /// <summary>
        ///     Move focus position rightward.
        /// </summary>
        /// <param name="autoReturn">Determines whether or not move to next column if reached end row.</param>
        public void MoveFocusRight(bool autoReturn = true)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            FocusPos = FindNextMovableCellRight(focusPos,
                RangeIsMergedCell(selectionRange) ? FixRange(RangePosition.EntireRange) : selectionRange,
                autoReturn);
        }

        /// <summary>
        ///     Move focus position downward.
        /// </summary>
        /// <param name="autoReturn">Determines whether or not move to next row if reached end column.</param>
        public void MoveFocusDown(bool autoReturn = true)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            FocusPos = FindNextMovableCellDown(focusPos,
                RangeIsMergedCell(selectionRange) ? FixRange(RangePosition.EntireRange) : selectionRange,
                autoReturn);
        }

        #region Move Utility

        private CellPosition FindNextMovableCellUp(CellPosition pos, int firstRow)
        {
            var row = pos.Row;

            // find next movable cell upward
            while (row > firstRow)
            {
                row--;

                var cell = cells[row, pos.Col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && row < cell.MergeEndPos.Row
                                 && row >= cell.MergeStartPos.Row)
                    continue;

                if (rows[row].InnerHeight > 0) break;
            }

            return new CellPosition(row, pos.Col);
        }

        private CellPosition FindNextMovableCellLeft(CellPosition pos, int firstCol)
        {
            var col = pos.Col;

            // find next movable cell leftward
            while (col > firstCol)
            {
                col--;

                var cell = cells[pos.Row, col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && col < cell.MergeEndPos.Col
                                 && col >= cell.MergeStartPos.Col)
                    continue;

                if (cols[col].InnerWidth > 0) break;
            }

            return new CellPosition(pos.Row, col);
        }

        private CellPosition FindNextMovableCellRight(CellPosition pos, RangePosition moveRange, bool autoReturn = true)
        {
            var col = pos.Col;

            var endCol = selectionRange.Cols > 1 ? selectionRange.EndCol : cols.Count - 1;

            if (col >= endCol)
            {
                var newpos = FindNextMovableCellDown(new CellPosition(pos.Row, moveRange.Col), moveRange, false);
                if (pos == newpos) return pos;

                pos = newpos;
            }

            var row = pos.Row;

            // find next movable cell rightward
            while (col < moveRange.EndCol)
            {
                col++;

                var cell = cells[row, col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && col <= cell.MergeEndPos.Col
                                 && col > cell.MergeStartPos.Col)
                    continue;

                if (cols[col].InnerWidth > 0) break;
            }

            return new CellPosition(pos.Row, col);
        }

        private CellPosition FindNextMovableCellDown(CellPosition pos, RangePosition moveRange, bool autoReturn = true)
        {
            var row = pos.Row;

            // find next movable cell downward
            while (row < moveRange.EndRow)
            {
                row++;

                var cell = cells[row, pos.Col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && row <= cell.MergeEndPos.Row
                                 && row > cell.MergeStartPos.Row)
                    continue;

                if (rows[row].InnerHeight > 0) break;
            }

            return new CellPosition(row, pos.Col);
        }

        #endregion // Move Utility

        /// <summary>
        ///     Move forward selection
        /// </summary>
        public void MoveSelectionForward()
        {
            if (SelectionMovedForward != null)
            {
                var arg = new SelectionMovedForwardEventArgs();
                SelectionMovedForward(this, arg);
                if (arg.IsCancelled) return;
            }

#if EX_SCRIPT
			var scriptReturn = RaiseScriptEvent("onnextfocus");
			if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
			{
				return;
			}
#endif

            switch (selectionForwardDirection)
            {
                case SelectionForwardDirection.Right:
                {
                    if (selEnd.Col < cols.Count - 1)
                    {
                        MoveSelectionRight();
                    }
                    else
                    {
                        if (selEnd.Row < rows.Count - 1)
                        {
                            selEnd.Col = 0;
                            MoveSelectionDown();
                        }
                    }
                }
                    break;

                case SelectionForwardDirection.Down:
                {
                    if (selEnd.Row < rows.Count - 1)
                    {
                        selEnd.Col = focusReturnColumn;

                        MoveSelectionDown();
                    }
                    else
                    {
                        if (selEnd.Col < cols.Count - 1)
                        {
                            selEnd.Row = 0;
                            MoveSelectionRight();
                        }
                    }
                }
                    break;
            }
        }

        /// <summary>
        ///     Move backward selection
        /// </summary>
        public void MoveSelectionBackward()
        {
            if (SelectionMovedBackward != null)
            {
                var arg = new SelectionMovedBackwardEventArgs();
                SelectionMovedBackward(this, arg);
                if (arg.IsCancelled) return;
            }

#if EX_SCRIPT
			var scriptReturn = RaiseScriptEvent("onpreviousfocus");
			if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
			{
				return;
			}
#endif

            switch (selectionForwardDirection)
            {
                case SelectionForwardDirection.Right:
                {
                    if (selEnd.Col > 0) MoveSelectionLeft();
                }
                    break;

                case SelectionForwardDirection.Down:
                {
                    if (selEnd.Row > 0) MoveSelectionUp();
                }
                    break;
            }
        }

        /// <summary>
        ///     Upward to move focus selection
        /// </summary>
        /// <param name="appendSelect">Decide whether or not perform an appending select (same as Shift key press down)</param>
        public void MoveSelectionUp(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var row = selEnd.Row;

            // downward to find next movable cell
            while (row > 0)
            {
                row--;

                var cell = cells[row, selEnd.Col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && row < cell.MergeEndPos.Row
                                 && row >= cell.MergeStartPos.Row)
                    continue;

                if (rows[row].InnerHeight > 0)
                {
                    MoveRangeSelection(selStart, new CellPosition(row, selEnd.Col), appendSelect);
                    break;
                }
            }
        }

        /// <summary>
        ///     Downward to move focus selection
        /// </summary>
        /// <param name="appendSelect">Decide whether or not perform an appending select (same as Shift key press down)</param>
        public void MoveSelectionDown(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var row = selEnd.Row;

            // downward to find next movable cell
            while (row < rows.Count - 1)
            {
                row++;

                var cell = cells[row, selEnd.Col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && row <= cell.MergeEndPos.Row
                                 && row > cell.MergeStartPos.Row)
                    continue;

                if (rows[row].InnerHeight > 0)
                {
                    MoveRangeSelection(selStart, new CellPosition(row, selEnd.Col), appendSelect);
                    break;
                }
            }
        }

        /// <summary>
        ///     Leftward to move focus selection
        /// </summary>
        /// <param name="appendSelect">Decide whether or not perform an appending select (same as Shift key press down)</param>
        public void MoveSelectionLeft(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var col = selEnd.Col;

            // downward to find next movable cell
            while (col > 0)
            {
                col--;

                var cell = cells[selEnd.Row, col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && col < cell.MergeEndPos.Col
                                 && col >= cell.MergeStartPos.Col)
                    continue;

                if (cols[col].InnerWidth > 0)
                {
                    //selEnd.Col = col;
                    MoveRangeSelection(selStart, new CellPosition(selEnd.Row, col), appendSelect);
                    break;
                }
            }
        }

        /// <summary>
        ///     Rightward to move focus selection
        /// </summary>
        /// <param name="appendSelect">Decide whether or not perform an appending select (same as Shift key press down)</param>
        public void MoveSelectionRight(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var col = selEnd.Col;

            // downward to find next movable cell
            while (col < cols.Count - 1)
            {
                col++;

                var cell = cells[selEnd.Row, col];

                if (cell != null && !cell.MergeEndPos.IsEmpty
                                 && col <= cell.MergeEndPos.Col
                                 && col > cell.MergeStartPos.Col)
                    continue;

                if (cols[col].InnerWidth > 0)
                {
                    MoveRangeSelection(selStart, new CellPosition(selEnd.Row, col), appendSelect);
                    break;
                }
            }
        }

        /// <summary>
        ///     Move selection to first cell of row or column which is specified by <code>rowOrColumn</code>
        /// </summary>
        /// <param name="rowOrColumn">specifies that move selection to first cell of row or column</param>
        /// <param name="appendSelect">Decide whether or not perform an appending select (same as Shift key press down)</param>
        public void MoveSelectionHome(RowOrColumn rowOrColumn, bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            //bool selectionChanged = false;
            var endpos = selEnd;

            if ((rowOrColumn & RowOrColumn.Row) == RowOrColumn.Row) endpos.Row = 0;

            if ((rowOrColumn & RowOrColumn.Column) == RowOrColumn.Column) endpos.Col = 0;

            if (endpos != selEnd) MoveRangeSelection(selStart, endpos, appendSelect);
        }

        /// <summary>
        ///     Move selection to last cell of row or column which is specified by <code>rowOrColumn</code>
        /// </summary>
        /// <param name="rowOrColumn">specifies that move selection to the cell of row or column</param>
        /// <param name="appendSelect">Determines that whether or not to expand the current selection.</param>
        public void MoveSelectionEnd(RowOrColumn rowOrColumn, bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            //bool selectionChanged = false;
            var endpos = selEnd;

            if ((rowOrColumn & RowOrColumn.Row) == RowOrColumn.Row) endpos.Row = rows.Count - 1;

            if ((rowOrColumn & RowOrColumn.Column) == RowOrColumn.Column) endpos.Col = cols.Count - 1;

            if (endpos != selEnd) MoveRangeSelection(selStart, endpos, appendSelect);
        }

        /// <summary>
        ///     Move selection to cell in next page vertically.
        /// </summary>
        /// <param name="appendSelect">
        ///     When this value is true, the selection will be expanded to the cell in next page rather than
        ///     moving it.
        /// </param>
        public void MoveSelectionPageDown(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var row = selEnd.Row;

            var nvc = ViewportController as NormalViewportController;

            if (nvc != null)
            {
                var viewport = nvc.FocusView as IViewport;

                if (viewport != null) row += Math.Max(viewport.VisibleRegion.Rows - 1, 1);
            }

            var pos = FixPos(new CellPosition(row, selEnd.Col));

            var cell = cells[pos.Row, pos.Col];

            if (cell != null)
            {
                cell = GetMergedCellOfRange(pos.Row, selEnd.Col);
                pos = cell.Position;
            }

            MoveRangeSelection(selStart, pos, appendSelect);
        }

        /// <summary>
        ///     Move selection to cell in previous page vertically.
        /// </summary>
        /// <param name="appendSelect">
        ///     When this value is true, the selection will be expanded to the cell in previous page rather
        ///     than moving it.
        /// </param>
        public void MoveSelectionPageUp(bool appendSelect = false)
        {
            if (selectionMode == WorksheetSelectionMode.None) return;

            var row = selEnd.Row;

            var nvc = ViewportController as NormalViewportController;

            if (nvc != null)
            {
                var viewport = nvc.FocusView as IViewport;

                if (viewport != null) row -= Math.Max(viewport.VisibleRegion.Rows - 1, 1);
            }

            var pos = FixPos(new CellPosition(row, selEnd.Col));

            var cell = cells[pos.Row, pos.Col];

            if (cell != null)
            {
                cell = GetMergedCellOfRange(pos.Row, selEnd.Col);
                pos = cell.Position;
            }

            MoveRangeSelection(selStart, pos, appendSelect);
        }

        #endregion // Keyboard Move

        #region Events

        /// <summary>
        ///     Event raised before selection range changing
        /// </summary>
        public event EventHandler<BeforeSelectionChangeEventArgs> BeforeSelectionRangeChange;

        /// <summary>
        ///     Event raised on focus-selection-range changed
        /// </summary>
        public event EventHandler<RangeEventArgs> SelectionRangeChanged;

        /// <summary>
        ///     Event raised on focus-selection-range is changing by mouse move
        /// </summary>
        public event EventHandler<RangeEventArgs> SelectionRangeChanging;

        /// <summary>
        ///     Event raised on Selection-Mode change
        /// </summary>
        public event EventHandler SelectionModeChanged;

        /// <summary>
        ///     Event raised on Selection-Style change
        /// </summary>
        public event EventHandler SelectionStyleChanged;

        /// <summary>
        ///     Event raised on SelectionForwardDirection change
        /// </summary>
        public event EventHandler SelectionForwardDirectionChanged;

        /// <summary>
        ///     Event raised when focus-selection move to next position
        /// </summary>
        public event EventHandler<SelectionMovedForwardEventArgs> SelectionMovedForward;

        /// <summary>
        ///     Event raised when focus-selection move to previous position
        /// </summary>
        public event EventHandler<SelectionMovedBackwardEventArgs> SelectionMovedBackward;

        #endregion // Events
    }
}