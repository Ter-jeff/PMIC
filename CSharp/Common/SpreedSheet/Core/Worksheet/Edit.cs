#define WPF

#if EX_SCRIPT
using unvell.ReoScript;
using unvell.ReoGrid.Script;
#endif // EX_SCRIPT;

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using System;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using SpreedSheet.View;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WPF

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        #region Editing Text

        /// <summary>
        ///     Get or set the current text in edit textbox of cell
        /// </summary>
        public string CellEditText
        {
            // TODO: move to control
            get { return controlAdapter.GetEditControlText(); }
            set { controlAdapter.SetEditControlText(value); }
        }

        #endregion // Editing Text

        #region StartEdit

        internal Cell CurrentEditingCell { get; set; }
        private string backupData;

        /// <summary>
        ///     Start to edit selected cell
        /// </summary>
        /// <returns>True if the editing operation has been started</returns>
        public bool StartEdit()
        {
            return selStart.IsEmpty ? false : StartEdit(focusPos);
        }

        /// <summary>
        ///     Start to edit selected cell
        /// </summary>
        /// <returns>True if the editing operation has been started</returns>
        public bool StartEdit(string newText)
        {
            return selStart.IsEmpty ? false : StartEdit(focusPos, newText);
        }

        /// <summary>
        ///     Start to edit specified cell
        /// </summary>
        /// <param name="pos">Position of specified cell</param>
        /// <returns>True if the editing operation has been started</returns>
        public bool StartEdit(CellPosition pos)
        {
            return StartEdit(pos.Row, pos.Col);
        }

        /// <summary>
        ///     Start to edit specified cell
        /// </summary>
        /// <param name="pos">Position of specified cell</param>
        /// <param name="newText">A text will be displayed in the edit field initially.</param>
        /// <returns>True if the editing operation has been started</returns>
        public bool StartEdit(CellPosition pos, string newText)
        {
            return StartEdit(pos.Row, pos.Col, newText);
        }

        /// <summary>
        ///     Start to edit specified cell
        /// </summary>
        /// <param name="row">Index of row of specified cell</param>
        /// <param name="col">Index of column of specified cell</param>
        /// <returns>True if the editing operation has been started</returns>
        public bool StartEdit(int row, int col)
        {
            if (row < 0 || col < 0 || row >= rows.Count || col >= cols.Count) return false;

            // if cell is part of merged cell
            if (!IsValidCell(row, col))
            {
                // find the merged cell
                var cell = GetMergedCellOfRange(row, col);

                // start edit on merged cell
                return StartEdit(cell);
            }

            return StartEdit(CreateAndGetCell(row, col));
        }

        /// <summary>
        ///     Start to edit specified cell.
        /// </summary>
        /// <param name="row">Index of row of specified cell.</param>
        /// <param name="col">Index of column of specified cell.</param>
        /// <param name="newText">A text displayed in the text field to be edited.</param>
        /// <returns>True if worksheet entered edit-mode successfully; Otherwise return false.</returns>
        public bool StartEdit(int row, int col, string newText)
        {
            if (row < 0 || col < 0 || row >= cells.RowCapacity || col >= cells.ColCapacity) return false;

            // if cell is part of merged cell
            if (!IsValidCell(row, col))
            {
                // find the merged cell
                var cell = GetMergedCellOfRange(row, col);

                // start edit on merged cell
                return StartEdit(cell, newText);
            }

            return StartEdit(CreateAndGetCell(row, col), newText);
        }

        internal bool StartEdit(Cell cell)
        {
            return StartEdit(cell, null);
        }

        internal bool StartEdit(Cell cell, string newText)
        {
            // abort if either spreadsheet or cell is readonly
            if (HasSettings(WorksheetSettings.Edit_Readonly)
                || cell == null || cell.IsReadOnly)
                return false;

            if (focusPos != cell.Position)
                FocusPos = cell.Position;
            else
                ScrollToCell(cell);

            string editText = null;

            if (newText == null)
            {
                if (!string.IsNullOrEmpty(cell.InnerFormula))
                    editText = "=" + cell.InnerFormula;
                else if (cell.InnerData is string)
                    editText = (string)cell.InnerData;
#if DRAWING
				else if (cell.InnerData is Drawing.RichText)
				{
					editText = ((Drawing.RichText)cell.InnerData).ToString();
				}
#endif // DRAWING
                else
                    editText = Convert.ToString(cell.InnerData);

                backupData = editText;
            }
            else
            {
                editText = newText;

                backupData = cell.DisplayText;
            }

            if (cell.DataFormat == CellDataFormatFlag.Percent
                && HasSettings(WorksheetSettings.Edit_FriendlyPercentInput))
            {
                double val;
                if (double.TryParse(editText, out val)) editText = (newText == null ? val * 100 : val) + "%";
            }

            if (BeforeCellEdit != null)
            {
                var arg = new CellBeforeEditEventArgs(cell)
                {
                    EditText = editText
                };

                BeforeCellEdit(this, arg);

                if (arg.IsCancelled) return false;

                editText = arg.EditText;
            }

#if EX_SCRIPT
			// v0.8.2: 'beforeCellEdit' renamed to 'onCellEdit'
			// v0.8.5: 'onCellEdit' renamed to 'oncelledit'
			object scriptReturn = RaiseScriptEvent("oncelledit", new RSCellObject(this, cell.InternalPos, cell));
			if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
			{
				return false;
			}
#endif

            if (cell.body != null)
            {
                var canContinue = cell.body.OnStartEdit();
                if (!canContinue) return false;
            }

            if (CurrentEditingCell != null) EndEdit(controlAdapter.GetEditControlText());

            CurrentEditingCell = cell;

            controlAdapter.SetEditControlText(editText);

            if (cell.DataFormat == CellDataFormatFlag.Percent && editText.EndsWith("%"))
                controlAdapter.SetEditControlCaretPos(editText.Length - 1);

            double x = 0;

            var width = (cell.Width - 1) * renderScaleFactor;

            var cellIndentSize = 0;

            //if ((cell.InnerStyle.Flag & PlainStyleFlag.Indent) == PlainStyleFlag.Indent)
            //{
            //indentSize = (int)Math.Round(cell.InnerStyle.Indent * this.indentSize * this.scaleFactor);
            //width -= indentSize;
            //}

#if WINFORM
			if (width < cell.TextBounds.Width) width = cell.TextBounds.Width;
#elif WPF
            // why + 6 ?
            if (width < cell.TextBounds.Width) width = cell.TextBounds.Width + 6;
#endif

            width--;
            //width = (width - 1);

            var scale = renderScaleFactor;

            #region Horizontal alignment

            switch (cell.RenderHorAlign)
            {
                default:
                case GridRenderHorAlign.Left:
                    controlAdapter.SetEditControlAlignment(GridHorAlign.Left);
                    x = cell.Left * scale + 1 + cellIndentSize;
                    break;

                case GridRenderHorAlign.Center:
                    controlAdapter.SetEditControlAlignment(GridHorAlign.Center);
                    x = cell.Left * scale + ((cell.Width - 1) * scale - 1 - width) / 2 + 1;
                    break;

                case GridRenderHorAlign.Right:
                    controlAdapter.SetEditControlAlignment(GridHorAlign.Right);
                    x = (cell.Right - 1) * scale - width - cellIndentSize;
                    break;
            }

            if (cell.InnerStyle.HAlign == GridHorAlign.DistributedIndent)
                controlAdapter.SetEditControlAlignment(GridHorAlign.Center);

            #endregion // Horizontal alignment

            var y = cell.Top * scale + 1;

            var activeViewport = ViewportController.FocusView as IViewport;

            var boxX = (int)Math.Round(x + ViewportController.FocusView.Left -
                                       (activeViewport == null ? 0 : activeViewport.ScrollViewLeft * scale));
            var boxY = (int)Math.Round(y + ViewportController.FocusView.Top -
                                       (activeViewport == null ? 0 : activeViewport.ScrollViewTop * scale));

            var height = (cell.Height - 1) * scale - 1;

            if (!cell.IsMergedCell && cell.InnerStyle.TextWrapMode != TextWrapMode.NoWrap)
                if (height < cell.TextBounds.Height)
                    height = cell.TextBounds.Height;

            var offsetHeight = 0; // (int)Math.Round(height);// (int)Math.Round(height + 2 - (cell.Height));

            if (offsetHeight > 0)
                switch (cell.InnerStyle.VAlign)
                {
                    case GridVerAlign.Top:
                        break;
                    default:
                    case GridVerAlign.Middle:
                        boxY -= offsetHeight / 2;
                        break;
                    case GridVerAlign.Bottom:
                        boxY -= offsetHeight;
                        break;
                }

            var rect = new Rectangle(boxX, boxY, width, height);

            controlAdapter.ShowEditControl(rect, cell);

            return true;
        }

        #endregion // StartEdit

        #region EndEdit

        /// <summary>
        ///     Check whether any cell current in edit mode
        /// </summary>
        /// <returns>true if any cell is editing</returns>
        public bool IsEditing
        {
            get { return CurrentEditingCell != null; }
        }

        /// <summary>
        ///     Get instance of current editing cell.
        /// </summary>
        public Cell EditingCell
        {
            get { return CurrentEditingCell; }
        }

        private bool endEditProcessing;

        /// <summary>
        ///     Force end current editing operation with the specified reason.
        /// </summary>
        /// <param name="reason">Ending Reason of editing operation</param>
        /// <returns>
        ///     True if currently in editing mode, and operation has been
        ///     finished successfully.
        /// </returns>
        public bool EndEdit(EndEditReason reason)
        {
            return EndEdit(reason == EndEditReason.NormalFinish ? controlAdapter.GetEditControlText() : null, reason);
        }

        /// <summary>
        ///     Force end current editing operation.
        ///     Uses specified data instead of the data of user edited.
        /// </summary>
        /// <param name="data">New data to be set to the edited cell</param>
        /// <returns>
        ///     True if currently in editing mode, and operation has been
        ///     finished successfully.
        /// </returns>
        public bool EndEdit(object data)
        {
            return EndEdit(data, EndEditReason.NormalFinish);
        }

        /// <summary>
        ///     Force end current editing operation with the specified reason.
        ///     Uses specified data instead of the data of user edited.
        /// </summary>
        /// <param name="data">New data to be set to the edited cell</param>
        /// <param name="reason">Ending Reason of editing operation</param>
        /// <returns>
        ///     True if currently in editing mode, and operation has been
        ///     finished successfully.
        /// </returns>
        public bool EndEdit(object data, EndEditReason reason)
        {
            if (CurrentEditingCell == null || endEditProcessing) return false;

            endEditProcessing = true;

            if (data == null) data = controlAdapter.GetEditControlText();

            if (AfterCellEdit != null)
            {
                var arg = new CellAfterEditEventArgs(CurrentEditingCell)
                {
                    EndReason = reason,
                    NewData = data
                };

                AfterCellEdit(this, arg);
                data = arg.NewData;
                reason = arg.EndReason;
            }

            switch (reason)
            {
                case EndEditReason.Cancel:
                    break;

                case EndEditReason.NormalFinish:
                    if (data is string)
                    {
                        var datastr = (string)data;

                        if (string.IsNullOrEmpty(datastr))
                            data = null;
                        else
                            // convert data into cell data format
                            switch (CurrentEditingCell.DataFormat)
                            {
                                case CellDataFormatFlag.Number:
                                case CellDataFormatFlag.Currency:
                                    double numericValue;
                                    if (double.TryParse(datastr, out numericValue)) data = numericValue;
                                    break;

                                case CellDataFormatFlag.Percent:
                                    if (datastr.EndsWith("%"))
                                    {
                                        double val;
                                        if (double.TryParse(datastr.Substring(0, datastr.Length - 1), out val))
                                            data = val / 100;
                                    }
                                    else if (datastr == "%")
                                    {
                                        data = null;
                                    }

                                    break;

                                case CellDataFormatFlag.DateTime:
                                {
                                    DateTime dt;
                                    if (DateTime.TryParse(datastr, out dt)) data = dt;
                                }
                                    break;
                            }
                    }

                    if (string.IsNullOrEmpty(backupData)) backupData = null;

                    var body = CurrentEditingCell.body;

                    if (body != null) data = body.OnEndEdit(data);

                    if (!Equals(data, backupData))
                        DoAction(new SetCellDataAction(CurrentEditingCell.InternalRow, CurrentEditingCell.InternalCol,
                            data));

                    break;
            }

            controlAdapter.HideEditControl();
            controlAdapter.Focus();
            CurrentEditingCell = null;

            endEditProcessing = false;

            return true;
        }

        #endregion // EndEdit

        #region Events

        /// <summary>
        ///     Event raised before cell changed to edit mode
        /// </summary>
        public event EventHandler<CellBeforeEditEventArgs> BeforeCellEdit;

        /// <summary>
        ///     Event raised after cell changed to edit mode
        /// </summary>
        public event EventHandler<CellAfterEditEventArgs> AfterCellEdit;

        /// <summary>
        ///     Event raised after input text changing
        /// </summary>
        public event EventHandler<CellEditTextChangingEventArgs> CellEditTextChanging;

        /// <summary>
        ///     Event raised after any characters is input
        /// </summary>
        public event EventHandler<CellEditCharInputEventArgs> CellEditCharInputed;

        internal string RaiseCellEditTextChanging(string text)
        {
            if (CellEditTextChanging == null) return text;

            var arg = new CellEditTextChangingEventArgs(CurrentEditingCell) { Text = text };
            CellEditTextChanging(this, arg);
            return arg.Text;
        }

        internal int RaiseCellEditCharInputed(int @char)
        {
            if (CellEditCharInputed == null) return @char;

            var arg = new CellEditCharInputEventArgs(CurrentEditingCell,
                CurrentEditingCell != null ? controlAdapter.GetEditControlText() : null,
                @char, controlAdapter.GetEditControlCaretPos(),
                controlAdapter.GetEditControlCaretLine());

            CellEditCharInputed(this, arg);

            return arg.InputChar;
        }

        #endregion // Events
    }
}