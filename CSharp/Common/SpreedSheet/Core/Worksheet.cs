#define WPF

#if WINFORM || ANDROID
using RGIntDouble = System.Int32;
using RGFloat = System.Single;
#elif WPF
using System;
using System.Diagnostics;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Utility;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
using RGIntDouble = System.Double;
#endif // WPF

namespace unvell.ReoGrid.Core
{
}

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        internal WorksheetRangeStyle RootStyle { get; set; }

        #region UpdateCellBounds

        private void UpdateCellBounds(Cell cell)
        {
#if DEBUG
            Debug.Assert(cell.Rowspan >= 1 && cell.Colspan >= 1);
#else
			if (cell.Rowspan < 1 || cell.Colspan < 1) return;
#endif
            cell.Bounds = GetRangeBounds(cell.InternalRow, cell.InternalCol, cell.Rowspan, cell.Colspan);
            UpdateCellTextBounds(cell);
            cell.UpdateContentBounds();
        }

        #endregion // UpdateCellBounds

        #region Set Style

        /// <summary>
        ///     Set styles to each cells inside specified range
        /// </summary>
        /// <param name="addressOrName">address or name to locate the cell or range on spreadsheet</param>
        /// <param name="style">styles to be set</param>
        /// <exception cref="InvalidAddressException">throw if specified address or name is illegal</exception>
        public void SetRangeStyles(string addressOrName, WorksheetRangeStyle style)
        {
            NamedRange namedRange;

            if (RangePosition.IsValidAddress(addressOrName))
                SetRangeStyles(new RangePosition(addressOrName), style);
            else if (registeredNamedRanges.TryGetValue(addressOrName, out namedRange))
                SetRangeStyles(namedRange, style);
            else
                throw new InvalidAddressException(addressOrName);
        }

        /// <summary>
        ///     Set styles to each cells inside specified range
        /// </summary>
        /// <param name="row">number of row of specified range</param>
        /// <param name="col">number of col of specified range</param>
        /// <param name="rows">number of rows inside specified range</param>
        /// <param name="cols">number of columns inside specified range</param>
        /// <param name="style">styles to be set</param>
        public void SetRangeStyles(int row, int col, int rows, int cols, WorksheetRangeStyle style)
        {
            SetRangeStyles(new RangePosition(row, col, rows, cols), style);
        }

        /// <summary>
        ///     Set styles to each cells inside specified range
        /// </summary>
        /// <param name="range">specified range to the styles</param>
        /// <param name="style">styles to be set</param>
        public void SetRangeStyles(RangePosition range, WorksheetRangeStyle style)
        {
            if (CurrentEditingCell != null) EndEdit(EndEditReason.NormalFinish);

            var fixedRange = FixRange(range);

            var r1 = fixedRange.Row;
            var c1 = fixedRange.Col;
            var r2 = fixedRange.EndRow;
            var c2 = fixedRange.EndCol;

            var isColStyle = fixedRange.Rows == rows.Count;
            var isRowStyle = fixedRange.Cols == cols.Count;
            var isRootStyle = isRowStyle && isColStyle;

            var isRange = !isColStyle && !isRowStyle;

            int maxRow, maxCol;
            if (isColStyle && r2 > (maxRow = MaxContentRow)) r2 = maxRow;
            if (isRowStyle && c2 > (maxCol = MaxContentCol)) c2 = maxCol;

            var pkind = StyleParentKind.Own;

            // update default styles
            if (isRootStyle)
            {
                #region All headers style updating

                StyleUtility.CopyStyle(style, RootStyle);

                // update styles which has been set into row headers
                for (var r = 0; r < rows.Count; r++)
                {
                    var rowHead = rows[r];

                    if (rowHead != null && rowHead.InnerStyle != null)
                        StyleUtility.CopyStyle(style, rowHead.InnerStyle);
                }

                // update styles which has been set into column headers
                for (var c = 0; c < cols.Count; c++)
                {
                    var colHead = cols[c];

                    //if (colHead != null && colHead.InnerStyle != null)
                    //{
                    //	unvell.ReoGrid.Utility.StyleUtility.CopyStyle(style, colHead.InnerStyle);
                    //}
                    if (colHead != null) colHead.InnerStyle = null;
                }

                #endregion

                pkind = StyleParentKind.Root;
            }
            else if (isRowStyle)
            {
                #region Rows in range updating

                // Rows in range updating
                for (var r = r1; r <= r2; r++)
                {
                    var rowHeader = rows[r];

                    if (rowHeader.InnerStyle == null)
                        rowHeader.InnerStyle = StyleUtility.CreateMergedStyle(style, RootStyle);
                    else
                        StyleUtility.CopyStyle(style, rowHeader.InnerStyle);
                }

                #endregion // Rows in range updating

                pkind = StyleParentKind.Row;
            }
            else if (isColStyle)
            {
                #region Columns in range updating

                // Columns in range updating
                for (var c = c1; c <= c2; c++)
                {
                    var colHeader = cols[c];

                    if (colHeader.InnerStyle == null)
                        colHeader.InnerStyle = StyleUtility.CreateMergedStyle(style, RootStyle);
                    else
                        StyleUtility.CopyStyle(style, colHeader.InnerStyle);
                }

                #endregion // Columns in range updating

                pkind = StyleParentKind.Col;
            }

            WorksheetRangeStyle rowStyle = null;
            WorksheetRangeStyle colStyle = null;

            // update cells
            for (var r = r1; r <= r2; r++)
            {
                rowStyle = null;

                for (var c = c1; c <= c2; c++)
                {
                    var cell = cells[r, c];
                    colStyle = null;

                    if (cell != null)
                    {
                        if (
                            cell.IsValidCell
                            &&

                            // if is a part of merged cell, check whether all rows or columns is selected
                            // if all rows or columns is selected, skip set styles
                            ((!isRowStyle && !isColStyle)
                             || (r1 <= cell.InternalRow && r2 >= cell.MergeEndPos.Row
                                                        && c1 <= cell.InternalCol && c2 >= cell.MergeEndPos.Col))
                        )
                        {
                            if (pkind == StyleParentKind.Row)
                            {
                                if (cell.StyleParentKind == StyleParentKind.Col)
                                {
                                    SetCellStyle(cell, style, StyleParentKind.Own);
                                }
                                else
                                {
                                    if (rowStyle == null) rowStyle = rows[r].InnerStyle;

                                    SetCellStyle(cell, style, pkind, rowStyle);
                                }
                            }
                            else if (pkind == StyleParentKind.Col)
                            {
                                if (colStyle == null) colStyle = cols[c].InnerStyle;
                                SetCellStyle(cell, style, pkind, colStyle);
                            }
                            else
                            {
                                SetCellStyle(cell, style, pkind, RootStyle);
                            }
                        }
                    }
                    else
                    {
                        // allow to create cells
                        if (isRange)
                        {
                            cell = CreateCell(r, c, false);
                            SetCellStyle(cell, style, StyleParentKind.Own);
                        }
                        // if full grid style then skip all null cells
                        else if (isRootStyle)
                        {
                        }
                        // if the column of cell has styles, compare to row style
                        else if (isColStyle)
                        {
                            if (rowStyle == null) rowStyle = rows[r].InnerStyle;

                            // if row has style, then create cell, else skip creating null cell
                            if (rowStyle != null)
                                // full column selected but the row of cell has also style,
                                // row style has the higher priority than the column style,
                                // so it is need to create instance of cell to 
                                // get highest priority for cell styles
                                SetCellStyle(CreateCell(r, c, false), style, StyleParentKind.Own);
                        }
                    }
                }
            }

            if (RangeStyleChanged != null) RangeStyleChanged(this, new RangeEventArgs(fixedRange));

            RequestInvalidate();
        }

        internal void SetCellStyleOwn(Cell cell, WorksheetRangeStyle style)
        {
            SetCellStyle(cell, style, StyleParentKind.Own);
        }

        internal void SetCellStyleOwn(CellPosition pos, WorksheetRangeStyle style)
        {
            SetCellStyleOwn(pos.Row, pos.Col, style);
        }

        /// <summary>
        ///     Set style to cell specified by row and col index
        /// </summary>
        /// <param name="row">index to row</param>
        /// <param name="col">index to col</param>
        /// <param name="style">style will be copied</param>
        internal void SetCellStyleOwn(int row, int col, WorksheetRangeStyle style)
        {
            SetCellStyle(CreateAndGetCell(row, col), style, StyleParentKind.Own);
        }

        private void SetCellStyle(Cell cell, WorksheetRangeStyle style,
            StyleParentKind parentKind, WorksheetRangeStyle parentStyle = null)
        {
            // do nothing if cell is a part of merged range
            if (cell.Rowspan == 0 || cell.Colspan == 0) return;

            if (cell.StyleParentKind == StyleParentKind.Own
                || parentKind == StyleParentKind.Own)
            {
                if (cell.StyleParentKind != StyleParentKind.Own) cell.CreateOwnStyle();

                StyleUtility.CopyStyle(style, cell.InnerStyle);

                // auto remove fill pattern when pattern color is empty
                if ((cell.InnerStyle.Flag & PlainStyleFlag.FillPattern) == PlainStyleFlag.FillPattern
                    && cell.InnerStyle.FillPatternColor.ToArgb() == 0)
                    cell.InnerStyle.Flag &= ~PlainStyleFlag.FillPattern;

                // auto remove background color when backcolor is empty
                if ((cell.InnerStyle.Flag & PlainStyleFlag.BackColor) == PlainStyleFlag.BackColor
                    && cell.InnerStyle.BackColor.ToArgb() == 0)
                    cell.InnerStyle.Flag &= ~PlainStyleFlag.BackColor;
            }
            else
            {
                cell.InnerStyle = parentStyle != null ? parentStyle : style;
                cell.StyleParentKind = parentKind;
            }

            // update render text align when data format changed
            StyleUtility.UpdateCellRenderAlign(this, cell);

            if (!string.IsNullOrEmpty(cell.DisplayText))
            {
                // when font changed, cell's scaled font need be updated.
                if (style.Flag.HasAny(PlainStyleFlag.FontAll))
                    // update cell font and text's bounds
                    UpdateCellFont(cell);
                // when font is not changed but alignment is changed, only update the bounds of text
                else if (style.Flag.HasAny(PlainStyleFlag.HorizontalAlign
                                           | PlainStyleFlag.VerticalAlign
                                           | PlainStyleFlag.TextWrap
                                           | PlainStyleFlag.Indent
                                           | PlainStyleFlag.RotationAngle))
                    UpdateCellTextBounds(cell);
#if WPF
                else if (style.Flag.Has(PlainStyleFlag.TextColor))
                    UpdateCellFont(cell, UpdateFontReason.TextColorChanged);
#endif // WPF
            }
            //else
            //{
            //	cell.RenderFont = null;
            //}

            // update cell bounds
            if (style.Flag.Has(PlainStyleFlag.Padding)) cell.UpdateContentBounds();

            // update cell body alignment
            if (cell.body != null && style.Flag.HasAny(PlainStyleFlag.AlignAll)) cell.body.OnBoundsChanged();
        }

        /// <summary>
        ///     Event raised on style of range changed
        /// </summary>
        public event EventHandler<RangeEventArgs> RangeStyleChanged;

        #endregion // Set Style

        #region Remove Style

        /// <summary>
        ///     Remove specified styles from a range specified by address or name
        /// </summary>
        /// <param name="addressOrName">Address or name to locate range from spreadsheet</param>
        /// <param name="flags">Styles specified by this flags to be removed</param>
        public void RemoveRangeStyles(string addressOrName, PlainStyleFlag flags)
        {
            if (RangePosition.IsValidAddress(addressOrName))
            {
                RemoveRangeStyles(new RangePosition(addressOrName), flags);
            }
            else
            {
                NamedRange namedRange;
                if (registeredNamedRanges.TryGetValue(addressOrName, out namedRange))
                    RemoveRangeStyles(namedRange, flags);
                else
                    throw new InvalidAddressException(addressOrName);
            }
        }

        /// <summary>
        ///     Remove specified styles from a specified range
        /// </summary>
        /// <param name="range">Range to be remove styles</param>
        /// <param name="flags">Styles specified by this flags to be removed</param>
        public void RemoveRangeStyles(RangePosition range, PlainStyleFlag flags)
        {
            var fixedRange = FixRange(range);

            var startRow = fixedRange.Row;
            var startCol = fixedRange.Col;
            var endRow = fixedRange.EndRow;
            var endCol = fixedRange.EndCol;

            var isFullColSelected = fixedRange.Rows == rows.Count;
            var isFullRowSelected = fixedRange.Cols == cols.Count;
            var isFullGridSelected = isFullRowSelected && isFullColSelected;

            var canCreateCell = !isFullColSelected && !isFullRowSelected;

            // update default styles
            if (isFullGridSelected)
            {
                RootStyle.Flag &= ~flags;

                // remote styles if it is already setted in full-row
                for (var r = 0; r < rows.Count; r++)
                {
                    var rowStyle = rows[r].InnerStyle;
                    if (rowStyle != null) rowStyle.Flag &= ~flags;
                }

                // remote styles if it is already setted in full-col
                for (var c = 0; c < cols.Count; c++)
                {
                    var colStyle = cols[c].InnerStyle;
                    if (colStyle != null) colStyle.Flag &= ~flags;
                }
            }
            else if (isFullRowSelected)
            {
                for (var r = startRow; r <= endRow; r++)
                {
                    var rowStyle = rows[r].InnerStyle;
                    if (rowStyle != null) rowStyle.Flag &= ~flags;
                }
            }
            else if (isFullColSelected)
            {
                for (var c = startCol; c <= endCol; c++)
                {
                    var colStyle = cols[c].InnerStyle;
                    if (colStyle != null) colStyle.Flag &= ~flags;
                }
            }

            for (var r = startRow; r <= endRow; r++)
            for (var c = startCol; c <= endCol;)
            {
                var cell = cells[r, c];

                if (cell == null)
                {
                    c++;
                }
                else if (cell.Rowspan == 1 && cell.Colspan == 1)
                {
                    RemoveCellStyle(cell, flags);
                    c++;
                }
                else if (cell.IsStartMergedCell
                         // only set merged cell if selection contains the merged entire range
                         && startRow <= cell.MergeStartPos.Row && endRow >= cell.MergeEndPos.Row
                         && startCol <= cell.MergeStartPos.Col && endCol >= cell.MergeEndPos.Col)
                {
                    RemoveCellStyle(cell, flags);
                    c += cell.Colspan;
                }
                else if (!cell.MergeStartPos.IsEmpty)
                {
                    c = cell.MergeEndPos.Col + 1;
                }
            }

            RequestInvalidate();
        }

        private void RemoveCellStyle(Cell cell, PlainStyleFlag flags)
        {
            // backup cell flags, copy the items from parent style by this flags
            var pFlag = cell.StyleParentKind;

            // cell style references to root style
            if (pFlag == StyleParentKind.Root)
            {
                var distinctedStyle = StyleUtility.CheckDistinctStyle(RootStyle, DefaultStyle);

                if (distinctedStyle == PlainStyleFlag.None)
                    // root style doesn't have own styles, no need to remove any styles
                    return;
            }

            // Parent style of cell
            WorksheetRangeStyle pStyle = null;
            var newPKind = StyleParentKind.Root;

            var rowhead = rows[cell.Row];
            var colhead = cols[cell.Column];

            // find parent style
            if (rowhead.InnerStyle != null)
            {
                pStyle = rowhead.InnerStyle;
                newPKind = StyleParentKind.Row;
            }
            else
            {
                if (colhead.InnerStyle != null)
                {
                    pStyle = colhead.InnerStyle;
                    newPKind = StyleParentKind.Col;
                }
                else
                {
                    pStyle = RootStyle;
                    newPKind = StyleParentKind.Root;
                }
            }

            if (pFlag != StyleParentKind.Own)
            {
                cell.InnerStyle = new WorksheetRangeStyle(pStyle);
                cell.InnerStyle.Flag &= ~flags;
            }
            else
            {
                cell.InnerStyle.Flag &= ~flags;

                // cell with own styles all have been removed
                // restore the cell to parent reference
                if (cell.InnerStyle.Flag == PlainStyleFlag.None)
                {
                    cell.InnerStyle = pStyle;
                    cell.StyleParentKind = newPKind;

                    return;
                }

                // remove style values
                if ((flags & PlainStyleFlag.BackColor) == PlainStyleFlag.BackColor)
                    cell.InnerStyle.BackColor = SolidColor.Transparent;
            }
            // remove style items by removing-flags

            switch (pFlag)
            {
                case StyleParentKind.Row:
                    if (colhead.InnerStyle != null)
                        pStyle = colhead.InnerStyle;
                    else
                        pStyle = RootStyle;
                    break;

                default:
                    pStyle = RootStyle;
                    break;
            }

            // copy style items from parent cell
            // copy all items by cellFlags in order to restore the style items from parent
            var newFlags = flags & pStyle.Flag;

            if (newFlags != PlainStyleFlag.None) StyleUtility.CopyStyle(pStyle, cell.InnerStyle, newFlags);

            // copy flags from parent style
            //cell.InnerStyle.Flag = pStyle.Flag;

            cell.StyleParentKind = StyleParentKind.Own;

            if ((flags & ( /*PlainStyleFlag.AlignAll  // may don't need this  |*/
                        PlainStyleFlag.TextWrap |
#if WINFORM || WPF || iOS
                        PlainStyleFlag.FontAll
#elif ANDROID
				PlainStyleFlag.FontName | PlainStyleFlag.FontStyleAll
#endif // ANDROID
                    )) > 0)
                cell.FontDirty = true;
        }

        #endregion // Remove Style

        #region Get Style

        /// <summary>
        ///     Get style of specified range.
        /// </summary>
        /// <param name="range">The range to get style.</param>
        /// <returns>Style info of specified range.</returns>
        public WorksheetRangeStyle GetRangeStyles(RangePosition range)
        {
            var fixedRange = FixRange(range);

            return GetCellStyles(range.StartPos);
        }

        internal object GetRangeStyle(int row, int col, int rows, int cols, PlainStyleFlag flag)
        {
            // TODO: return range's style instead of cell 
            return GetCellStyleItem(row, col, flag);
        }

        /// <summary>
        ///     Get style from cell by specified position.
        /// </summary>
        /// <param name="address">Address to locate a cell to get its style.</param>
        /// <returns>Style set of cell retrieved from specified position.</returns>
        public WorksheetRangeStyle GetCellStyles(string address)
        {
            if (!CellPosition.IsValidAddress(address)) throw new InvalidAddressException(address);

            return GetCellStyles(new CellPosition(address));
        }

        /// <summary>
        ///     Get style of single cell.
        /// </summary>
        /// <param name="pos">Position of cell to get.</param>
        /// <returns>Style of cell in the specified position.</returns>
        public WorksheetRangeStyle GetCellStyles(CellPosition pos)
        {
            return GetCellStyles(pos.Row, pos.Col);
        }

        /// <summary>
        ///     Get style of specified cell without creating its instance.
        /// </summary>
        /// <param name="row">Index of row of specified cell.</param>
        /// <param name="col">Index of column of specified cell.</param>
        /// <returns>Style of cell from specified position.</returns>
        public WorksheetRangeStyle GetCellStyles(int row, int col)
        {
            var cell = cells[row, col];
            var pKind = StyleParentKind.Own;
            if (cell == null)
                return StyleUtility.FindCellParentStyle(this, row, col, out pKind);
            return new WorksheetRangeStyle(cell.InnerStyle);
        }

        /// <summary>
        ///     Get single style item from specified cell
        /// </summary>
        /// <param name="row">Zero-based number of row</param>
        /// <param name="col">Zero-based number of column</param>
        /// <param name="flag">Specified style item to be get</param>
        /// <returns>Style item value</returns>
        public object GetCellStyleItem(int row, int col, PlainStyleFlag flag)
        {
            var cell = cells[row, col];

            var pKind = StyleParentKind.Own;

            var style = cell == null ? StyleUtility.FindCellParentStyle(this, row, col, out pKind) : cell.InnerStyle;

            return StyleUtility.GetStyleItem(style, flag);
        }

        #endregion

        #region Update Font & Text

        internal void UpdateCellFont(Cell cell, UpdateFontReason reason = UpdateFontReason.FontChanged)
        {
            UpdateCellRenderFont(null, cell, DrawMode.View, reason);
        }

        internal void UpdateCellRenderFont(IRenderer ir, Cell cell, DrawMode drawMode, UpdateFontReason reason)
        {
            if (controlAdapter == null || cell.InnerStyle == null) return;

            // cell doesn't contain any text, clear font dirty flag and return
            if (string.IsNullOrEmpty(cell.InnerDisplay))
                // can't use below sentence, that makes RenderFont property null due to unknown reasons
                //cell.FontDirty = false;
                return;

#if DRAWING
			// rich text object doesn't need update font
			if (!(cell.Data is Drawing.RichText))
			{
#endif // DRAWING

            if (ir == null) ir = controlAdapter.Renderer;

            ir.UpdateCellRenderFont(cell, reason);

#if DRAWING
			}
#endif // DRAWING

            cell.FontDirty = false;

            UpdateCellTextBounds(ir, cell, drawMode, reason);
        }

        /// <summary>
        ///     Update Cell Text Bounds for View/Edit mode
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="updateRowHeight"></param>
        internal void UpdateCellTextBounds(Cell cell)
        {
            if (cell.FontDirty)
                UpdateCellFont(cell);
            else
                UpdateCellTextBounds(null, cell, DrawMode.View, UpdateFontReason.FontChanged);
        }

        internal void UpdateCellTextBounds(IRenderer ig, Cell cell, DrawMode drawMode, UpdateFontReason reason)
        {
            UpdateCellTextBounds(ig, cell, drawMode, renderScaleFactor, reason);
        }

        /// <summary>
        ///     Update cell text bounds.
        ///     need to call this method when content of cell is changed, contains styles like align, font, etc.
        ///     if cell's display property is null, this method does nothing.
        /// </summary>
        /// <param name="ig">The graphics device used to calculate bounds. Null to use default graphic device.</param>
        /// <param name="cell">The target cell will be updated.</param>
        /// <param name="drawMode">Draw mode</param>
        /// <param name="scaleFactor">Scale factor of current worksheet</param>
        internal void UpdateCellTextBounds(IRenderer ig, Cell cell, DrawMode drawMode, double scaleFactor,
            UpdateFontReason reason)
        {
            if (cell == null || string.IsNullOrEmpty(cell.DisplayText)) return;

            if (ig == null && controlAdapter != null) ig = controlAdapter.Renderer;

            if (ig == null) return;

            Size oldSize;
            Size size;

#if DRAWING
			if (cell.Data is Drawing.RichText)
			{
				var rt = (Drawing.RichText)cell.Data;

				oldSize = rt.TextSize;

				rt.TextWrap = cell.InnerStyle.TextWrapMode;
				rt.DefaultHorizontalAlignment = cell.Style.HAlign;
				rt.VerticalAlignment = cell.Style.VAlign;

				rt.Size = new Size(cell.Width - cell.InnerStyle.Indent, cell.Height);

				size = rt.TextSize;
				return;
			}
			else
			{
#endif // DRAWING

            oldSize = cell.TextBounds.Size;

            #region Plain Text Measure Size

            size = ig.MeasureCellText(cell, drawMode, scaleFactor);

            if (size.Width <= 0 || size.Height <= 0) return;

            // FIXME: get incorrect size if CJK fonts
            size.Width += 2;
            size.Height += 1;

            #endregion // Plain Text Measure Size

            var cellBounds = cell.Bounds;

            var cellWidth = cellBounds.Width * scaleFactor;

#if WINFORM
				if (cell.InnerStyle.HAlign == ReoGridHorAlign.DistributedIndent)
				{
					size.Width--;

					if (drawMode == DrawMode.View)
					{
						cell.DistributedIndentSpacing =
 ((cellWidth - size.Width - 3) / (cell.DisplayText.Length - 1)) - 1;
						if (cell.DistributedIndentSpacing < 0) cell.DistributedIndentSpacing = 0;
					}
					else
					{
						cell.DistributedIndentSpacingPrint =
 ((cellWidth - size.Width - 3) / (cell.DisplayText.Length - 1)) - 1;
						if (cell.DistributedIndentSpacingPrint < 0) cell.DistributedIndentSpacingPrint = 0;
					}

					cell.RenderHorAlign = ReoGridRenderHorAlign.Center;
					if (size.Width < cellWidth - 1) size.Width = (float)(Math.Round(cellWidth - 1));
				}

#elif WPF

            if (cell.InnerStyle.TextWrapMode != TextWrapMode.NoWrap) cell.formattedText.MaxTextWidth = cellWidth;

#endif // WPF

            #region Update Text Size Cache

            double x = 0;
            double y = 0;

            float indent = cell.InnerStyle.Indent;

            switch (cell.RenderHorAlign)
            {
                default:
                case GridRenderHorAlign.Left:
                    x = cellBounds.Left * scaleFactor + 2 + indent * IndentSize;
                    break;

                case GridRenderHorAlign.Center:
                    x = cellBounds.Left * scaleFactor + cellWidth / 2 - size.Width / 2;
                    break;

                case GridRenderHorAlign.Right:
                    x = cellBounds.Right * scaleFactor - 3 - size.Width - indent * IndentSize;
                    break;
            }

            switch (cell.InnerStyle.VAlign)
            {
                case GridVerAlign.Top:
                    y = cellBounds.Top * scaleFactor + 1;
                    break;

                case GridVerAlign.Middle:
                    y = cellBounds.Top * scaleFactor + cellBounds.Height * scaleFactor / 2 - size.Height / 2;
                    break;

                default:
                case GridVerAlign.General:
                case GridVerAlign.Bottom:
                    y = cellBounds.Bottom * scaleFactor - 1 - size.Height;
                    break;
            }

            switch (drawMode)
            {
                default:
                case DrawMode.View:
                    cell.TextBounds = new Rectangle(x, y, size.Width, size.Height);
                    break;

                case DrawMode.Preview:
                case DrawMode.Print:
                    cell.PrintTextBounds = new Rectangle(x, y, size.Width, size.Height);
                    break;
            }

            #endregion // Update Text Size Cache

#if DRAWING
			}
#endif // DRAWING

            if (drawMode == DrawMode.View
                && reason != UpdateFontReason.ScaleChanged)
            {
                if (size.Height > oldSize.Height
                    && settings.Has(WorksheetSettings.Edit_AutoExpandRowHeight))
                {
                    var rowHeader = rows[cell.Row];

                    if (rowHeader.IsVisible && rowHeader.IsAutoHeight) cell.ExpandRowHeight();
                }

                if (size.Width > oldSize.Width
                    && settings.Has(WorksheetSettings.Edit_AutoExpandColumnWidth))
                {
                    var colHeader = cols[cell.Column];

                    if (colHeader.IsVisible && colHeader.IsAutoWidth) cell.ExpandColumnWidth();
                }
            }
        }

        /// <summary>
        ///     Make the text of cells in specified range larger or smaller.
        /// </summary>
        /// <param name="range">The spcified range.</param>
        /// <param name="stepHandler">Iterator callback to handle how to make text larger or smaller.</param>
        public void StepRangeFont(RangePosition range, Func<float, float> stepHandler)
        {
            var fixedRange = FixRange(range);

            var enableAdjustRowHeight = settings.Has(WorksheetSettings.Edit_AllowAdjustRowHeight
                                                     | WorksheetSettings.Edit_AutoExpandRowHeight);

            RowHeader rowHeader = null;

            IterateCells(fixedRange, (r, c, cell) =>
            {
                cell.CreateOwnStyle();

                var newSize = stepHandler(cell.InnerStyle.FontSize);

                if (enableAdjustRowHeight && newSize > cell.InnerStyle.FontSize)
                {
                    cell.InnerStyle.FontSize = newSize;

                    if (rowHeader == null || rowHeader.Index != r)
                    {
                        rowHeader = rows[r];

                        if (rowHeader.IsAutoHeight) cell.ExpandRowHeight();
                    }
                }
                else
                {
                    cell.InnerStyle.FontSize = newSize;
                }

                cell.FontDirty = true;

                return true;
            });

            RequestInvalidate();
        }

        #endregion // Update Font & Text
    }

    #region Alignment

    #endregion // Alignment

    #region enum PlainStyleFlag

    #endregion // PlainStyleFlag

    #region enum TextWrapMode

    #endregion // TextWrapMode

    #region Padding

    #endregion

    #region WorksheetRangeStyle

    #endregion // ReoGridStyleObject

    #region ReferenceStyle

    #endregion // ReferenceStyle

    #region ReferenceRangeStyle

    #endregion // ReferenceRangeStyle

    #region ColumnHeaderStyle

    #endregion // ColumnHeaderStyle

    #region RowHeaderStyle

    #endregion // RowHeaderStyle

    #region ReferenceCellStyle

    #endregion // ReferenceCellStyle
}