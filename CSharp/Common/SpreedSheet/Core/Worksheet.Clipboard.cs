using System;
using System.Linq;
using System.Text;
using System.Windows;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Interface;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        private static readonly string ClipBoardDataFormatIdentify = "{CB3BE3D1-2BF9-4fa6-9B35-374F6A0412CE}";

        private RangePosition currentCopingRange = RangePosition.Empty;

        public string StringifyRange(string addressOrName)
        {
            if (RangePosition.IsValidAddress(addressOrName)) return StringifyRange(new RangePosition(addressOrName));

            NamedRange namedRange;
            if (registeredNamedRanges.TryGetValue(addressOrName, out namedRange))
                return StringifyRange(namedRange);
            throw new InvalidAddressException(addressOrName);
        }

        /// <summary>
        ///     Convert all data from specified range to a tabbed string.
        /// </summary>
        /// <param name="range">The range to be converted.</param>
        /// <returns>Tabbed string contains all data converted from specified range.</returns>
        public string StringifyRange(RangePosition range)
        {
            var erow = range.EndRow;
            var ecol = range.EndCol;

            // copy plain text
            var sb = new StringBuilder();

            var isFirst = true;
            for (var r = range.Row; r <= erow; r++)
            {
                if (isFirst) isFirst = false;
                else sb.Append('\n');

                var isFirst2 = true;
                for (var c = range.Col; c <= ecol; c++)
                {
                    if (isFirst2) isFirst2 = false;
                    else sb.Append('\t');

                    var cell = cells[r, c];
                    if (cell != null)
                    {
                        var text = cell.DisplayText;

                        if (!string.IsNullOrEmpty(text))
                        {
                            if (text.Contains('\n')) text = string.Format("\"{0}\"", text);

                            sb.Append(text);
                        }
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        ///     Paste data from tabbed string into worksheet.
        /// </summary>
        /// <param name="address">Start cell position to be filled.</param>
        /// <param name="content">Data to be pasted.</param>
        /// <returns>Range position that indicates the actually filled range.</returns>
        public RangePosition PasteFromString(string address, string content)
        {
            if (!CellPosition.IsValidAddress(address)) throw new InvalidAddressException(address);

            return PasteFromString(new CellPosition(address), content);
        }

        /// <summary>
        ///     Paste data from tabbed string into worksheet.
        /// </summary>
        /// <param name="startPos">Start position to fill data.</param>
        /// <param name="content">Tabbed string to be pasted.</param>
        /// <returns>Range position that indicates the actually filled range.</returns>
        public RangePosition PasteFromString(CellPosition startPos, string content)
        {
            //int rows = 0, cols = 0;

            //string[] lines = content.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            //for (int r = 0; r < lines.Length; r++)
            //{
            //	string line = lines[r];
            //	if (line.EndsWith("\n")) line = line.Substring(0, line.Length - 1);
            //	//line = line.Trim();

            //	if (line.Length > 0)
            //	{
            //		string[] tabs = line.Split('\t');
            //		cols = Math.Max(cols, tabs.Length);

            //		for (int c = 0; c < tabs.Length; c++)
            //		{
            //			int toRow = startPos.Row + r;
            //			int toCol = startPos.Col + c;

            //			if (!this.IsValidCell(toRow, toCol))
            //			{
            //				throw new RangeIntersectionException(new RangePosition(toRow, toCol, 1, 1));
            //			}

            //			string text = tabs[c];

            //			if (text.StartsWith("\"") && text.EndsWith("\""))
            //			{
            //				text = text.Substring(1, text.Length - 2);
            //			}

            //			SetCellData(toRow, toCol, text);
            //		}

            //		rows++;
            //	}
            //}

            var parsedData = RGUtility.ParseTabbedString(content);

            var rows = parsedData.GetLength(0);
            var cols = parsedData.GetLength(1);

            var range = new RangePosition(startPos.Row, startPos.Col, rows, cols);

            SetRangeData(range, parsedData);

            return range;
        }

        #region Copy

        /// <summary>
        ///     Copy data and put into Clipboard.
        /// </summary>
        public bool Copy()
        {
            if (IsEditing)
            {
                controlAdapter.EditControlCopy();
            }
            else
            {
                controlAdapter.ChangeCursor(CursorStyle.Busy);

                try
                {
                    if (BeforeCopy != null)
                    {
                        var evtArg = new BeforeRangeOperationEventArgs(selectionRange);
                        BeforeCopy(this, evtArg);
                        if (evtArg.IsCancelled) return false;
                    }

                    // highlight current copy range
                    currentCopingRange = selectionRange;

                    var data = new DataObject();
                    data.SetData(ClipBoardDataFormatIdentify,
                        GetPartialGrid(currentCopingRange, PartialGridCopyFlag.All, ExPartialGridCopyFlag.None, true));

                    var text = StringifyRange(currentCopingRange);
                    if (!string.IsNullOrEmpty(text)) data.SetText(text);

                    // set object data into clipboard
                    Clipboard.SetDataObject(data);

                    if (AfterCopy != null) AfterCopy(this, new RangeEventArgs(selectionRange));
                }
                catch (Exception ex)
                {
                    NotifyExceptionHappen(ex);
                    return false;
                }
                finally
                {
                    controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
                }
            }

            return true;
        }

        #endregion

        #region Cut

        /// <summary>
        ///     Copy any remove anything from selected range into Clipboard.
        /// </summary>
        /// <param name="byAction">
        ///     Indicates whether or not perform the cut operation by using an action, which makes the operation
        ///     can be undone. Default is true.
        /// </param>
        /// <returns></returns>
        public bool Cut(bool byAction = true)
        {
            if (IsEditing)
            {
                controlAdapter.EditControlCut();
            }
            else
            {
                if (!Copy()) return false;

                if (BeforeCut != null)
                {
                    var evtArg = new BeforeRangeOperationEventArgs(selectionRange);

                    BeforeCut(this, evtArg);

                    if (evtArg.IsCancelled) return false;
                }

#if EX_SCRIPT
				object scriptReturn = RaiseScriptEvent("oncut");
				if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
				{
					return false;
				}
#endif

                if (!HasSettings(WorksheetSettings.Edit_Readonly))
                {
                    var data = Clipboard.GetDataObject() as DataObject;
                    var partialGrid = data.GetData(ClipBoardDataFormatIdentify) as PartialGrid;

                    var startRow = selectionRange.Row;
                    var startCol = selectionRange.Col;

                    var rows = partialGrid.Rows;
                    var cols = partialGrid.Columns;

                    var range = new RangePosition(startRow, startCol, rows, cols);

                    if (byAction)
                    {
                        DoAction(new CutRangeAction(range, partialGrid));
                    }
                    else
                    {
                        DeleteRangeData(range, true);
                        RemoveRangeStyles(range, PlainStyleFlag.All);
                        RemoveRangeBorders(range, BorderPositions.All);
                    }
                }

                if (AfterCut != null) AfterCut(this, new RangeEventArgs(selectionRange));
            }

            return true;
        }

        #endregion

        #region Paste

        /// <summary>
        ///     Copy data from Clipboard and put on grid.
        ///     Currently ReoGrid supports the following types of source from the clipboard.
        ///     - Data from another ReoGrid instance
        ///     - Plain/Unicode Text from any Windows Applications
        ///     - Tabbed Plain/Unicode Data from Excel or similar applications
        ///     When data copied from another ReoGrid instance, and the destination range
        ///     is bigger than the source, ReoGrid will try to repeat putting data to fill
        ///     the destination range entirely.
        ///     Todo: Copy border and cell style from Excel.
        /// </summary>
        public bool Paste()
        {
            if (IsEditing)
            {
                controlAdapter.EditControlPaste();
            }
            else
            {
                // Paste method will always perform action to do paste

                // do nothing if in readonly mode
                if (HasSettings(WorksheetSettings.Edit_Readonly)
                    // or selection is empty
                    || selectionRange.IsEmpty)
                    return false;

                try
                {
                    controlAdapter.ChangeCursor(CursorStyle.Busy);

                    PartialGrid partialGrid = null;
                    string clipboardText = null;

                    var data = Clipboard.GetDataObject() as DataObject;
                    if (data != null)
                    {
                        partialGrid = data.GetData(ClipBoardDataFormatIdentify) as PartialGrid;

                        if (data.ContainsText()) clipboardText = data.GetText();
                    }

                    if (partialGrid != null)
                    {
                        #region Partial Grid Pasting

                        var startRow = selectionRange.Row;
                        var startCol = selectionRange.Col;

                        var rows = partialGrid.Rows;
                        var cols = partialGrid.Columns;

                        var rowRepeat = 1;
                        var colRepeat = 1;

                        if (selectionRange.Rows % partialGrid.Rows == 0)
                        {
                            rows = selectionRange.Rows;
                            rowRepeat = selectionRange.Rows / partialGrid.Rows;
                        }

                        if (selectionRange.Cols % partialGrid.Columns == 0)
                        {
                            cols = selectionRange.Cols;
                            colRepeat = selectionRange.Cols / partialGrid.Columns;
                        }

                        var targetRange = new RangePosition(startRow, startCol, rows, cols);

                        if (!RaiseBeforePasteEvent(targetRange)) return false;

                        if (targetRange.EndRow >= this.rows.Count
                            || targetRange.EndCol >= this.cols.Count)
                            // TODO: paste range overflow
                            // need to notify user-code to handle this 
                            return false;

                        // check whether the range to be pasted contains readonly cell
                        if (CheckRangeReadonly(targetRange))
                        {
                            NotifyExceptionHappen(
                                new OperationOnReadonlyCellException("specified range contains readonly cell"));
                            return false;
                        }

                        // check any intersected merge-range in partial grid 
                        // 
                        var cancelPerformPaste = false;

                        if (partialGrid.Cells != null)
                            try
                            {
                                #region Check repeated intersected ranges

                                for (var rr = 0; rr < rowRepeat; rr++)
                                for (var cc = 0; cc < colRepeat; cc++)
                                    partialGrid.Cells.Iterate((row, col, cell) =>
                                    {
                                        if (cell.IsMergedCell)
                                            for (var r = startRow;
                                                 r < cell.MergeEndPos.Row - cell.InternalRow + startRow + 1;
                                                 r++)
                                            for (var c = startCol;
                                                 c < cell.MergeEndPos.Col - cell.InternalCol + startCol + 1;
                                                 c++)
                                            {
                                                var tr = r + rr * partialGrid.Rows;
                                                var tc = c + cc * partialGrid.Columns;

                                                var existedCell = cells[tr, tc];

                                                if (existedCell != null)
                                                {
                                                    if (
                                                        // cell is a part of merged cell
                                                        (existedCell.Rowspan == 0 && existedCell.Colspan == 0)
                                                        // cell is merged cell
                                                        || existedCell.IsMergedCell)
                                                        throw new RangeIntersectionException(selectionRange);
                                                    // cell is readonly
                                                    if (existedCell.IsReadOnly)
                                                        throw new CellDataReadonlyException(cell.InternalPos);
                                                }
                                            }

                                        return Math.Min(cell.Colspan, (short)1);
                                    });

                                #endregion

                                // Check repeated intersected ranges
                            }
                            catch (Exception ex)
                            {
                                cancelPerformPaste = true;

                                // raise event to notify user-code there is error happened during paste operation
                                if (OnPasteError != null)
                                    OnPasteError(this, new RangeOperationErrorEventArgs(selectionRange, ex));
                            }

                        if (!cancelPerformPaste)
                            DoAction(new SetPartialGridAction(new RangePosition(
                                startRow, startCol, rows, cols), partialGrid));

                        #endregion // Partial Grid Pasting
                    }
                    else if (!string.IsNullOrEmpty(clipboardText))
                    {
                        #region Plain Text Pasting

                        var arrayData = RGUtility.ParseTabbedString(clipboardText);

                        var rows = Math.Max(selectionRange.Rows, arrayData.GetLength(0));
                        var cols = Math.Max(selectionRange.Cols, arrayData.GetLength(1));

                        var targetRange = new RangePosition(selectionRange.Row, selectionRange.Col, rows, cols);
                        if (!RaiseBeforePasteEvent(targetRange)) return false;

                        if (controlAdapter != null)
                        {
                            var actionSupportedControl = controlAdapter.ControlInstance as IActionControl;

                            if (actionSupportedControl != null)
                                actionSupportedControl.DoAction(this, new SetRangeDataAction(targetRange, arrayData));
                        }

                        #endregion // Plain Text Pasting
                    }
                }
                catch (Exception ex)
                {
                    // raise event to notify user-code there is error happened during paste operation
                    //if (OnPasteError != null)
                    //{
                    //	OnPasteError(this, new RangeOperationErrorEventArgs(selectionRange, ex));
                    //}
                    NotifyExceptionHappen(ex);
                }
                finally
                {
                    controlAdapter.ChangeCursor(CursorStyle.Selection);

                    RequestInvalidate();
                }

                if (AfterPaste != null) AfterPaste(this, new RangeEventArgs(selectionRange));
            }

            return true;
        }

        private bool RaiseBeforePasteEvent(RangePosition range)
        {
            if (BeforePaste != null)
            {
                var evtArg = new BeforeRangeOperationEventArgs(range);
                BeforePaste(this, evtArg);
                if (evtArg.IsCancelled) return false;
            }

#if EX_SCRIPT
			object scriptReturn = RaiseScriptEvent("onpaste", new RSRangeObject(this, range));
			if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
			{
				return false;
			}
#endif // EX_SCRIPT

            return true;
        }

        #endregion

        #region Checks

        /// <summary>
        ///     Determine whether the selected range can be copied.
        /// </summary>
        /// <returns>True if the selected range can be copied.</returns>
        public bool CanCopy()
        {
            //TODO
            return true;
        }

        /// <summary>
        ///     Determine whether the selected range can be cutted.
        /// </summary>
        /// <returns>True if the selected range can be cutted.</returns>
        public bool CanCut()
        {
            //TODO
            return true;
        }

        /// <summary>
        ///     Determine whether the data contained in Clipboard can be pasted into grid control.
        /// </summary>
        /// <returns>True if the data contained in Clipboard can be pasted</returns>
        public bool CanPaste()
        {
            //TODO
            return true;
        }

        #endregion

        #region Events

        /// <summary>
        ///     Before a range will be pasted from Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforePaste;

        /// <summary>
        ///     When a range has been pasted into grid
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterPaste;

        /// <summary>
        ///     When an error happened during perform paste
        /// </summary>
        [Obsolete("use SheetControl.ErrorHappened instead")]
        public event EventHandler<RangeOperationErrorEventArgs> OnPasteError;

        /// <summary>
        ///     Before a range to be copied into Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforeCopy;

        /// <summary>
        ///     When a range has been copied into Clipboard
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterCopy;

        /// <summary>
        ///     Before a range to be moved into Clipboard
        /// </summary>
        public event EventHandler<BeforeRangeOperationEventArgs> BeforeCut;

        /// <summary>
        ///     After a range to be moved into Clipboard
        /// </summary>
        public event EventHandler<RangeEventArgs> AfterCut;

        #endregion
    }
}