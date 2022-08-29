using System;
using System.Collections.Generic;
using System.Diagnostics;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.Common;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Utility;

namespace SpreedSheet.View
{
    internal class CellsViewport : Viewport, IRangeSelectableView
    {
        public CellsViewport(IViewportController vc)
            : base(vc)
        {
        }

        #region Draw

        #region DrawView

        public override void DrawView(CellDrawingContext dc)
        {
            if (Sheet.rows.Count <= 0 || Sheet.cols.Count < 0) return;

            // view mode
            if (Sheet.HasSettings(WorksheetSettings.View_ShowGridLine)
                && dc.DrawMode == DrawMode.View
                // zoom < 40% will not display grid lines
                && ScaleFactor >= 0.4f)
                DrawGuideLines(dc);

            DrawContent(dc);

            DrawSelection(dc);
        }

        #endregion // DrawView

        internal void DrawContent(CellDrawingContext dc)
        {
            #region Cells

            var toRow = VisibleRegion.EndRow + (dc.FullCellClip ? 0 : 1);
            var toCol = VisibleRegion.EndCol + (dc.FullCellClip ? 0 : 1);

            #region Background-only Cells

            var drawedCells = new List<Cell>(5);

            for (var r = VisibleRegion.StartRow; r < toRow; r++)
            {
                var rowHead = Sheet.rows[r];
                if (rowHead.InnerHeight <= 0) continue;

                for (var c = VisibleRegion.StartCol; c < toCol;)
                {
                    var cell = Sheet.cells[r, c];

                    if (cell == null)
                    {
                        DrawCellBackground(dc, r, c, cell);
                        c++;
                    }
                    else if (string.IsNullOrEmpty(cell.DisplayText))
                    {
                        if (cell.Rowspan == 1 && cell.Colspan == 1)
                        {
                            DrawCellBackground(dc, r, c, cell);
                            c++;
                        }
                        else if (
                            cell.IsStartMergedCell
                            || (cell.IsEndMergedCell
                                && !VisibleRegion.Contains(cell.MergeStartPos)))
                        {
                            DrawCellBackground(dc, cell.MergeStartPos.Row, cell.MergeStartPos.Col,
                                Sheet.GetCell(cell.MergeStartPos));
#if DEBUG
                            Debug.Assert(cell.MergeEndPos.Col >= c);
#endif // DEBUG
                            c = cell.MergeEndPos.Col + 1;
                        }
                        // merged cell is outside of visible region should also to be drawn
                        else if (
                            //cell.InternalRow == visibleRegion.startRow
                            //&& cell.InternalCol == visibleRegion.startCol && 
                            cell.Rowspan == 0 || (cell.Colspan == 0
                                                  && ((cell.MergeStartPos.Row < VisibleRegion.StartRow
                                                       && cell.MergeEndPos.Row > VisibleRegion.EndRow)
                                                      || (cell.MergeStartPos.Col < VisibleRegion.StartCol
                                                          && cell.MergeEndPos.Col > VisibleRegion.EndCol))))
                        {
                            var mergedStartCell = Sheet.GetCell(cell.MergeStartPos);

                            if (!drawedCells.Contains(mergedStartCell))
                            {
                                DrawCellBackground(dc, mergedStartCell.Row, mergedStartCell.Column, mergedStartCell);
                                drawedCells.Add(mergedStartCell);
                            }

#if DEBUG
                            Debug.Assert(cell.MergeEndPos.Col >= c);
#endif // DEBUG

                            c = cell.MergeEndPos.Col + 1;
                        }
                        else
                        {
                            c++;
                        }
                    }
                    else
                    {
                        c++;
                    }
                }
            }

            #endregion // Background-only Cells

            #region Display Text Cells

            drawedCells.Clear();

            for (var r = VisibleRegion.StartRow; r < toRow && r <= Sheet.cells.MaxRow; r++)
            {
                var rowHead = Sheet.rows[r];
                if (rowHead.InnerHeight <= 0) continue;

                for (var c = VisibleRegion.StartCol; c < toCol && c <= Sheet.cells.MaxCol;)
                {
                    var cell = Sheet.cells[r, c];

                    // draw cell onyl when cell's instance existing
                    // and bounds of cell must be > 1 (minimum is 1, including one pixel border)
                    if (cell != null && cell.Width > 1 && cell.Height > 1)
                    {
                        var hasContent = !string.IsNullOrEmpty(cell.DisplayText) || cell.body != null;

                        // single cell
                        if (cell.Rowspan == 1 && cell.Colspan == 1 && hasContent)
                        {
                            DrawCell(dc, cell);
                            c++;
                        }

                        // merged cell start
                        else if (cell.IsStartMergedCell && hasContent)
                        {
                            DrawCell(dc, cell);
                            c = cell.MergeEndPos.Col + 1;
                        }

                        // merged cell end
                        else if (cell.IsEndMergedCell
                                 && !VisibleRegion.Contains(cell.MergeStartPos)
                                 // don't check hasContent because it is the current cell,
                                 // we should check and draw merged start cell
                                 //&& hasContent
                                )
                        {
                            var mergedStartCell = Sheet.GetCell(cell.MergeStartPos);

                            if (!string.IsNullOrEmpty(mergedStartCell.DisplayText) || mergedStartCell.body != null)
                                DrawCell(dc, Sheet.GetCell(cell.MergeStartPos));
                            c = cell.MergeEndPos.Col + 1;
                        }

                        // merged cell is outside of visible region should also to be drawn
                        else if (
                            //cell.InternalRow == visibleRegion.startRow
                            //&& cell.InternalCol == visibleRegion.startCol &&
                            (cell.MergeStartPos.Row < VisibleRegion.StartRow
                             && cell.MergeEndPos.Row > VisibleRegion.EndRow)
                            || (cell.MergeStartPos.Col < VisibleRegion.StartCol
                                && cell.MergeEndPos.Col > VisibleRegion.EndCol))
                        {
                            var mergedStartCell = Sheet.GetCell(cell.MergeStartPos);

                            if (!drawedCells.Contains(mergedStartCell))
                            {
                                if (!string.IsNullOrEmpty(mergedStartCell.DisplayText) || mergedStartCell.body != null)
                                    DrawCell(dc, mergedStartCell);

                                drawedCells.Add(mergedStartCell);
                            }

                            c = cell.MergeEndPos.Col + 1;
                        }
                        else
                        {
                            c++;
                        }
                    }
                    else
                    {
                        c++;
                    }
                }
            }

            #endregion // Display Text Cells

            #endregion // Cells

#if DEBUG
            var sw = new Stopwatch();
            sw.Reset();
            sw.Start();
#endif // DEBUG

            #region Vertical Borders

            var rightColBoundary = VisibleRegion.EndCol + (dc.FullCellClip ? 0 : 1);

            for (var c = VisibleRegion.StartCol; c <= rightColBoundary; c++)
            {
                var x = c == Sheet.cols.Count ? Sheet.cols[c - 1].Right : Sheet.cols[c].Left;

                if (c < Sheet.cols.Count)
                {
                    // skip invisible vertical borders
                    var colHeader = Sheet.cols[c];
                    if (!colHeader.IsVisible) continue;
                }

                for (var r = VisibleRegion.StartRow; r <= VisibleRegion.EndRow;)
                {
                    var y = r == Sheet.rows.Count ? Sheet.rows[r - 1].Bottom : Sheet.rows[r].Top;

                    var cellBorder = Sheet.vBorders[r, c];
                    if (cellBorder != null && cellBorder.Span > 0 && cellBorder.Style != null)
                    {
                        var endRow = r + Math.Min(cellBorder.Span - 1, VisibleRegion.EndRow);

                        if (dc.FullCellClip && endRow >= VisibleRegion.EndRow - 1) endRow = VisibleRegion.EndRow - 1;

                        var y2 = Sheet.rows[endRow].Bottom;

                        BorderPainter.Instance.DrawLine(dc.Graphics.PlatformGraphics, x * ScaleFactor, y * ScaleFactor,
                            x * ScaleFactor, y2 * ScaleFactor, cellBorder.Style);

                        r += cellBorder.Span;
                    }
                    else
                    {
                        r++;
                    }
                }
            }

            #endregion

            #region Horizontal Borders

            var rightRowBoundary = VisibleRegion.EndRow + (dc.FullCellClip ? 0 : 1);

            for (var r = VisibleRegion.StartRow; r <= rightRowBoundary; r++)
            {
                if (r < Sheet.rows.Count)
                {
                    // skip invisible horizontal borders
                    var rowHeader = Sheet.rows[r];
                    if (!rowHeader.IsVisible) continue;
                }

                var y = r == Sheet.rows.Count ? Sheet.rows[r - 1].Bottom : Sheet.rows[r].Top;

                for (var c = VisibleRegion.StartCol; c <= VisibleRegion.EndCol;)
                {
                    var x = c == Sheet.cols.Count ? Sheet.cols[c - 1].Right : Sheet.cols[c].Left;

                    var cellBorder = Sheet.hBorders[r, c];
                    if (cellBorder != null && cellBorder.Span > 0 && cellBorder.Style != null)
                    {
                        var endCol = c + Math.Min(cellBorder.Span - 1, VisibleRegion.EndCol);

                        if (dc.FullCellClip && endCol >= VisibleRegion.EndCol - 1) endCol = VisibleRegion.EndCol - 1;

                        var x2 = Sheet.cols[endCol].Right;

                        BorderPainter.Instance.DrawLine(dc.Graphics.PlatformGraphics, x * ScaleFactor, y * ScaleFactor,
                            x2 * ScaleFactor, y * ScaleFactor, cellBorder.Style);

                        c += cellBorder.Span;
                    }
                    else
                    {
                        c++;
                    }
                }
            }

            #endregion

#if DEBUG
            sw.Stop();
            if (sw.ElapsedMilliseconds > 1000) Debug.WriteLine($"draw border ({sw.ElapsedMilliseconds} ms.)");
#endif // DEBUG

            #region View Mode Visible

            if (dc.DrawMode == DrawMode.View)
            {
                #region Print Breaks

#if PRINT
				if (this.sheet.HasSettings(WorksheetSettings.View_ShowPageBreaks)
					&& this.sheet.pageBreakRows != null && this.sheet.pageBreakCols != null
					&& this.sheet.pageBreakRows.Count > 0 && this.sheet.pageBreakCols.Count > 0)
				{
					RGFloat minX = this.sheet.cols[this.sheet.pageBreakCols[0]].Left * this.ScaleFactor;
					RGFloat minY = this.sheet.rows[this.sheet.pageBreakRows[0]].Top * this.ScaleFactor;
					RGFloat maxX =
 this.sheet.cols[this.sheet.pageBreakCols[this.sheet.pageBreakCols.Count - 1] - 1].Right * this.ScaleFactor;
					RGFloat maxY =
 this.sheet.rows[this.sheet.pageBreakRows[this.sheet.pageBreakRows.Count - 1] - 1].Bottom * this.ScaleFactor;

					foreach (int row in this.sheet.pageBreakRows)
					{
						RGFloat y =
 (row >= this.sheet.rows.Count ? this.sheet.rows[row - 1].Bottom : this.sheet.rows[row].Top) * this.ScaleFactor;

						bool isUserPageSplitter =
 this.sheet.userPageBreakRows != null && this.sheet.userPageBreakRows.Contains(row);

						dc.Graphics.DrawLine(Math.Max(this.ScrollViewLeft * this.ScaleFactor, minX), y,
							Math.Min(this.ScrollViewLeft * this.ScaleFactor + bounds.Width, maxX), y,
							SolidColor.Blue, 2f, isUserPageSplitter ? LineStyles.Solid : LineStyles.Dash);
					}

					foreach (int col in this.sheet.pageBreakCols)
					{
						RGFloat x =
 (col >= this.sheet.cols.Count ? this.sheet.cols[col - 1].Right : this.sheet.cols[col].Left) * this.ScaleFactor;

						bool isUserPageSplitter =
 this.sheet.userPageBreakCols != null && this.sheet.userPageBreakCols.Contains(col);

						dc.Graphics.DrawLine(x, Math.Max(this.ScrollViewTop * this.ScaleFactor, minY),
							x, Math.Min(this.ScrollViewTop * this.ScaleFactor + bounds.Height, maxY),
							SolidColor.Blue, 2f, isUserPageSplitter ? LineStyles.Solid : LineStyles.Dash);
					}
				}
#endif // PRINT

                #endregion // Print Breaks

                #region Break Lines Adjusting

                if (Sheet.pageBreakAdjustCol > -1 && Sheet.pageBreakAdjustFocusIndex > -1)
                {
                    double x;

                    if (Sheet.pageBreakAdjustFocusIndex < Sheet.cols.Count)
                        x = Sheet.cols[Sheet.pageBreakAdjustFocusIndex].Left;
                    else
                        x = Sheet.cols[Sheet.cols.Count - 1].Right;

                    x *= ScaleFactor;

                    dc.Graphics.FillRectangle(HatchStyles.Percent50, SolidColor.Gray, SolidColor.Transparent,
                        x - 1, ScrollViewTop * ScaleFactor, 3, ScrollViewTop + Height);
                }

                if (Sheet.pageBreakAdjustRow > -1 && Sheet.pageBreakAdjustFocusIndex > -1)
                {
                    double y;

                    if (Sheet.pageBreakAdjustFocusIndex < Sheet.rows.Count)
                        y = Sheet.rows[Sheet.pageBreakAdjustFocusIndex].Top;
                    else
                        y = Sheet.rows[Sheet.rows.Count - 1].Bottom;

                    y *= ScaleFactor;

                    dc.Graphics.FillRectangle(HatchStyles.Percent50, SolidColor.Gray, SolidColor.Transparent,
                        ScrollViewLeft * ScaleFactor, y - 1, ScrollViewLeft + Width, 3);
                }

                #endregion // Break Lines Adjusting

                #region Highlight & Focus Ranges

                if (Sheet.highlightRanges != null)
                    foreach (var range in Sheet.highlightRanges)
                        // is visible?
                        if (range.HighlightColor.A > 0 && VisibleRegion.IsOverlay(range))
                            DrawHighlightRange(dc, range);

                var focusHr = Sheet.focusHighlightRange;

                if (focusHr != null && focusHr.HighlightColor.A > 0
                                    && VisibleRegion.IsOverlay(focusHr))
                {
                    var rect = GetScaledAndClippedRangeRect(this, focusHr.StartPos, focusHr.EndPos, 1f);
                    rect.Inflate(-1, -1);

                    dc.Renderer.DrawRunningFocusRect(rect.X, rect.Y, rect.Right, rect.Bottom,
                        focusHr.HighlightColor, focusHr.RunnerOffset);

                    focusHr.RunnerOffset += 2;

                    if (focusHr.RunnerOffset > 9) focusHr.RunnerOffset = 0;
                }

                #endregion // Highlight & Focus Ranges
            }

            #endregion // View Mode Visible

            #region Trace Precedents & Dependents

#if FORMULA
			if (sheet.traceDependentArrows != null && sheet.traceDependentArrows.Count > 0)
			{
				var r = dc.Renderer;

				RGFloat ellipseSize = 4 * this.ScaleFactor;
				RGFloat halfOfEllipse = ellipseSize / 2 + 1;

				r.BeginCappedLine(LineCapStyles.Ellipse, new Size(ellipseSize - 1, ellipseSize - 1),
					 LineCapStyles.Arrow, new Size(halfOfEllipse, ellipseSize), SolidColor.Blue, 1);

				foreach (var fromCell in sheet.traceDependentArrows.Keys)
				{
					var lines = sheet.traceDependentArrows[fromCell];

					foreach (var pl in lines)
					{
						if (visibleRegion.Contains(fromCell.InternalPos)
							&& visibleRegion.Contains(pl.InternalPos))
						{
							Point startPoint = GetScaledTracePoint(fromCell.InternalPos);
							Point endPoint = GetScaledTracePoint(pl.InternalPos);

							r.DrawCappedLine(startPoint.X, startPoint.Y, endPoint.X, endPoint.Y);
						}
					}
				}

				r.EndCappedLine();
			}
#endif // FORMULA

            #endregion // Trace Precedents & Dependents
        }

        #region Clip Utility

        private Point GetScaledTracePoint(CellPosition startPos)
        {
            var startBounds = Sheet.GetCellBounds(startPos);

            var startPoint = startBounds.Location;
            startPoint.X += Math.Min(startBounds.Width, 30) / 2;
            startPoint.Y += Math.Min(startBounds.Height, 30) / 2;
            startPoint.X *= ScaleFactor;
            startPoint.Y *= ScaleFactor;

            return startPoint;
        }

        internal static Rectangle GetScaledAndClippedRangeRect(IViewport view, CellPosition startPos,
            CellPosition endPos, float borderWidth)
        {
            var sheet = view.ViewportController.Worksheet;
            var rangeRect = sheet.GetRangeBounds(startPos, endPos);

            var rowHead = sheet.rows[startPos.Row];
            var colHead = sheet.cols[startPos.Col];
            var toRowHead = sheet.rows[endPos.Row];
            var toColHead = sheet.cols[endPos.Col];

            var width = toColHead.Right - colHead.Left;
            var height = toRowHead.Bottom - rowHead.Top;

            var scaledRangeRect = new Rectangle(
                colHead.Left * view.ScaleFactor,
                rowHead.Top * view.ScaleFactor,
                width * view.ScaleFactor,
                height * view.ScaleFactor);

            return GetClippedRangeRect(view, scaledRangeRect, borderWidth);
        }

        private static Rectangle GetClippedRangeRect(IViewport view, Rectangle scaledRangeRect, float borderWidth)
        {
            var scaledViewTop = view.ScrollViewTop * view.ScaleFactor;
            var scaledViewLeft = view.ScrollViewLeft * view.ScaleFactor;

            var viewBottom = view.Height + scaledViewTop + borderWidth; // 3: max select range border overflow
            var viewRight = view.Width + scaledViewLeft + borderWidth;

            // top
            if (scaledRangeRect.Y < scaledViewTop - borderWidth)
            {
                var h = scaledRangeRect.Height - scaledViewTop + scaledRangeRect.Y + borderWidth;
                if (h < 0) h = 0;
                scaledRangeRect.Height = h;
                scaledRangeRect.Y = scaledViewTop - borderWidth;
            }

            // left
            if (scaledRangeRect.X < scaledViewLeft - borderWidth)
            {
                var w = scaledRangeRect.Width - scaledViewLeft + scaledRangeRect.X + borderWidth;
                if (w < 0) w = 0;
                scaledRangeRect.Width = w;
                scaledRangeRect.X = scaledViewLeft - borderWidth;
            }

            // bottom
            if (scaledRangeRect.Bottom > viewBottom)
            {
                var h = viewBottom - scaledRangeRect.Y;
                if (h < 0) h = 0;
                scaledRangeRect.Height = h;
            }

            // right
            if (scaledRangeRect.Right > viewRight)
            {
                var w = viewRight - scaledRangeRect.X;
                if (w < 0) w = 0;
                scaledRangeRect.Width = w;
            }

            return scaledRangeRect;
        }

        #endregion // Clip Utility

        #region Draw Highlight Range

        private void DrawHighlightRange(CellDrawingContext dc, HighlightRange range)
        {
            var g = dc.Graphics;

            var color = range.HighlightColor;
            var weight = range.Hover ? 2f : 1f;

            // convert to view rectangle
            var scaledRange = Sheet.GetScaledRangeBounds(range);
            var clippedRange = GetClippedRangeRect(this, scaledRange, weight);

            g.DrawRectangle(clippedRange, color, weight, LineStyles.Solid);
            g.FillRectangle(scaledRange.X - 1, scaledRange.Y - 1, 5, 5, color);
            g.FillRectangle(scaledRange.Right - 3, scaledRange.Y - 1, 5, 5, color);
            g.FillRectangle(scaledRange.X - 1, scaledRange.Bottom - 3, 5, 5, color);
            g.FillRectangle(scaledRange.Right - 3, scaledRange.Bottom - 3, 5, 5, color);
        }

        #endregion // Draw Highlight Range

        #region Draw Guide Lines

        private void DrawGuideLines(CellDrawingContext dc)
        {
            var render = dc.Renderer;

            var endRow = VisibleRegion.EndRow + (dc.FullCellClip ? 0 : 1);
            var endCol = VisibleRegion.EndCol + (dc.FullCellClip ? 0 : 1);

            render.BeginDrawLine(1, Sheet.controlAdapter.ControlStyle.Colors[ControlAppearanceColors.GridLine]);

            #region Horizontal line

            // horizontal line
            for (var r = VisibleRegion.StartRow; r <= endRow; r++)
            {
                float y = r >= Sheet.rows.Count ? Sheet.rows[Sheet.rows.Count - 1].Bottom : Sheet.rows[r].Top;
                var scaledY = y * ScaleFactor;

                for (var c = VisibleRegion.StartCol; c < endCol; c++)
                {
                    // skip horizontal border - line start

                    var x = Sheet.cols[c].Left;
                    var x2 = x; // sheet.cols[c].Right;

                    // skip horizontal border - line end
                    while (c < endCol)
                    {
                        var cellBorder = Sheet.hBorders[r, c];

                        if (cellBorder != null && cellBorder.Span >= 0) break;

                        if (r > 0)
                        {
                            var cell = Sheet.cells[r, c];

                            if (cell != null && cell.InnerStyle.BackColor.A > 0) break;
                        }

                        c++;
                    }

                    x2 = c == 0 ? x : Sheet.cols[c - 1].Right;
                    render.DrawLine(x * ScaleFactor, scaledY, x2 * ScaleFactor, scaledY);
                }
            }

            #endregion // Horizontal line

            #region Vertical line

            // vertical line
            for (var c = VisibleRegion.StartCol; c <= endCol; c++)
            {
                float x = c == Sheet.cols.Count ? Sheet.cols[c - 1].Right : Sheet.cols[c].Left;
                var scaledX = x * ScaleFactor;

                for (var r = VisibleRegion.StartRow; r < endRow; r++)
                {
                    var y = Sheet.rows[r].Top;
                    var y2 = y; // sheet.rows[r].Bottom;

                    while (r < endRow)
                    {
                        var cellBorder = Sheet.vBorders[r, c];

                        if (cellBorder != null && cellBorder.Span >= 0) break;

                        if (c > 0)
                        {
                            var cell = Sheet.cells[r, c];
                            if (cell != null && cell.InnerStyle.BackColor.A > 0) break;
                        }

                        r++;
                    }

                    y2 = r == 0 ? y : Sheet.rows[r - 1].Bottom;
                    render.DrawLine(scaledX, y * ScaleFactor, scaledX, y2 * ScaleFactor);
                }
            }

            #endregion // Vertical line

            render.EndDrawLine();
        }

        #endregion // Draw Gridlines

        #region Draw Cells

        #region DrawCell Entry

        private void DrawCell(CellDrawingContext dc, Cell cell)
        {
            if (cell == null) return;

            if (cell.IsMergedCell && (cell.Width <= 1 || cell.Height <= 1)) return;

            if (cell.body != null)
            {
                dc.Cell = cell;

                var g = dc.Graphics;

                g.PushTransform();

                if (ScaleFactor != 1f) g.ScaleTransform(ScaleFactor, ScaleFactor);

                g.TranslateTransform(dc.Cell.Left, dc.Cell.Top);

                cell.body.OnPaint(dc);

                g.PopTransform();
            }
            else
            {
                if (!string.IsNullOrEmpty(cell.DisplayText))
                {
                    DrawCellBackground(dc, cell.InternalRow, cell.InternalCol, cell);

                    DrawCellText(dc, cell);
                }
            }
        }

        #endregion // DrawCell Entry

        #region DrawCell Text

        internal void DrawCellText(CellDrawingContext dc, Cell cell)
        {
            var g = dc.Graphics;

            #region Plain Text

            #region Determine text color

            SolidColor textColor;
            if (!cell.RenderColor.IsTransparent)
                // render color, used to render negative number, specified by data formatter
                textColor = cell.RenderColor;
            else if (cell.InnerStyle.HasStyle(PlainStyleFlag.TextColor))
                // cell text color, specified by SetRangeStyle
                textColor = cell.InnerStyle.TextColor;
            // default cell text color
            else if (!Sheet.controlAdapter.ControlStyle.TryGetColor(ControlAppearanceColors.GridText, out textColor))
                // default built-in text
                textColor = SolidColor.Black;

            if (cell.FontDirty) Sheet.UpdateCellFont(cell);

            #endregion

            #region Determine clip region

            var cellScaledWidth = cell.Width * ScaleFactor;
            double cellScaledHeight = (float)Math.Floor(cell.Height * ScaleFactor) - 1;

            var clipRect = new Rectangle(ScrollViewLeft * ScaleFactor, cell.Top * ScaleFactor, Width, cellScaledHeight);

            var needWidthClip = cell.IsMergedCell ||
                                cell.InnerStyle.TextWrapMode == TextWrapMode.WordBreak ||
                                dc.AllowCellClip;

            if (!needWidthClip)
            {
                if (cell.InternalCol < Sheet.cols.Count - 1)
                    if (cell.RenderHorAlign == GridRenderHorAlign.Left
                        || cell.RenderHorAlign == GridRenderHorAlign.Center)
                    {
                        var move = 1;
                        while (!needWidthClip)
                        {
                            if (cell.TextBounds.Right < cell.Right)
                                break;
                            if (cell.InternalCol + move > Sheet.MaxContentCol)
                                break;
                            var nextCell = Sheet.cells[cell.InternalRow, cell.InternalCol + move];
                            if (!(nextCell == null || string.IsNullOrEmpty(nextCell.DisplayText)))
                            {
                                needWidthClip = true;
                                clipRect = cell.Bounds;
                                clipRect.X *= ScaleFactor;
                                clipRect.Y *= ScaleFactor;
                                clipRect.Width = cellScaledWidth * move;
                                clipRect.Height = cellScaledHeight;

                                if (move != 1)
                                {
                                    var selectionBorderWidth = Sheet.controlAdapter.ControlStyle.SelectionBorderWidth;
                                    var rectangle = new Rectangle(cell.TextBounds.Left,
                                        cell.Top + selectionBorderWidth,
                                        cell.TextBounds.Width - 2 * selectionBorderWidth,
                                        cell.Height - 2 * selectionBorderWidth);
                                    g.DrawAndFillRectangle(rectangle, SolidColor.Red, SolidColor.Red);
                                }
                            }

                            move++;
                        }
                    }

                if (!needWidthClip
                    && cell.InternalCol > 0
                    && (cell.RenderHorAlign == GridRenderHorAlign.Right
                        || cell.RenderHorAlign == GridRenderHorAlign.Center))
                {
                    var prevCell = Sheet.cells[cell.InternalRow, cell.InternalCol - 1];
                    needWidthClip = prevCell != null
                                    && prevCell.TextBounds.Left < cell.Left
                                    && !string.IsNullOrEmpty(prevCell.DisplayText);
                }
            }

            if (!needWidthClip) needWidthClip = cell.TextBoundsHeight > cellScaledHeight;
            //var needWidthClip = true;
            //clipRect = cell.Bounds;
            //clipRect.X *= ScaleFactor;
            //clipRect.Y *= ScaleFactor;
            //clipRect.Width = cellScaledWidth;
            //clipRect.Height = cellScaledHeight;
            if (needWidthClip) g.PushClip(clipRect);

            #endregion

            dc.Renderer.DrawCellText(cell, textColor, dc.DrawMode, ScaleFactor);

            if (needWidthClip)
                dc.Graphics.PopClip();

            #endregion
        }

        #endregion // DrawCell Text

        #region DrawCell Background

        internal void DrawCellBackground(CellDrawingContext dc, int row, int col, Cell cell, bool refPosition = false)
        {
            WorksheetRangeStyle style;

            if (cell == null)
            {
                var pKind = StyleParentKind.Own;
                style = StyleUtility.FindCellParentStyle(Sheet, row, col, out pKind);
            }
            else
            {
                style = cell.InnerStyle;
            }

            if (style.BackColor.A > 0)
            {
                var startPos = new CellPosition(row, col);

                var rect = cell == null
                    ? GetScaledAndClippedRangeRect(this, startPos, startPos, 1)
                    : GetScaledAndClippedRangeRect(this, startPos,
                        new CellPosition(row + cell.Rowspan - 1, col + cell.Colspan - 1), 1);

                if (cell != null && refPosition) rect.Location = new Point(0, 0);

                if (rect.Width > 0 && rect.Height > 0)
                {
                    var g = dc.Graphics;

                    if (style.FillPatternColor.A > 0)
                        g.FillRectangle(style.FillPatternStyle, style.FillPatternColor, style.BackColor, rect);
                    else
                        g.FillRectangle(rect, style.BackColor);
                }
            }
        }

        #endregion // DrawCell Background

        #endregion // Draw Cells

        #region Draw Selection

        private void DrawSelection(CellDrawingContext dc)
        {
            // selection
            if (!Sheet.SelectionRange.IsEmpty
                && dc.DrawMode == DrawMode.View
                && Sheet.SelectionStyle != WorksheetSelectionStyle.None)
            {
                var g = dc.Graphics;
                var controlStyle = Sheet.workbook.controlAdapter.ControlStyle;

                var selectionBorderWidth = controlStyle.SelectionBorderWidth;

                var scaledSelectionRect = GetScaledAndClippedRangeRect(this,
                    Sheet.SelectionRange.StartPos, Sheet.SelectionRange.EndPos, selectionBorderWidth);

                if (scaledSelectionRect.Width > 0 || scaledSelectionRect.Height > 0)
                {
                    var selectionFillColor = controlStyle.Colors[ControlAppearanceColors.SelectionFill];

                    if (Sheet.SelectionStyle == WorksheetSelectionStyle.Default)
                    {
                        var range = Sheet.GetRangeIfMergedCell(Sheet.focusPos);
                        var scaledFocusPosRect = GetScaledAndClippedRangeRect(this, range.StartPos, range.EndPos, 0);
                        var selectionBorderColor = controlStyle.Colors[ControlAppearanceColors.SelectionBorder];
                        if (!Sheet.SelectionRange.IsSingleCell)
                            g.FillRectangle(scaledSelectionRect, selectionFillColor);
                        if (selectionBorderColor.A > 0)
                            g.DrawRectangle(scaledSelectionRect, selectionBorderColor, selectionBorderWidth,
                                LineStyles.Solid);
                    }
                    else if (Sheet.SelectionStyle == WorksheetSelectionStyle.FocusRect)
                    {
                        g.DrawRectangle(scaledSelectionRect, SolidColor.Black, 1, LineStyles.Dot);
                    }

                    if (Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToFillSerial))
                    {
                        var sheetBackColor = controlStyle.Colors[ControlAppearanceColors.GridBackground];

                        var thumbRect = new Rectangle(scaledSelectionRect.Right - selectionBorderWidth,
                            scaledSelectionRect.Bottom - selectionBorderWidth,
                            selectionBorderWidth + 2, selectionBorderWidth + 2);

                        g.DrawRectangle(thumbRect, sheetBackColor);
                    }
                }
            }
        }

        #endregion // Draw Selection

        #endregion // Draw

        #region Mouse

        public override bool OnMouseDown(Point location, MouseButtons buttons)
        {
            var isProcessed = false;
            if (!isProcessed
                && Sheet.selectionMode != WorksheetSelectionMode.None
                && !Sheet.HasSettings(WorksheetSettings.Edit_Readonly))
            {
                if (!isProcessed
                    && Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToFillSerial))
                {
                    #region Hit Selection Drag

                    if (SelectDragCornerHitTest(Sheet, location))
                    {
                        Sheet.operationStatus = OperationStatus.DragSelectionFillSerial;
                        Sheet.lastMouseMoving = location;
                        Sheet.draggingSelectionRange = Sheet.selectionRange;
                        Sheet.focusMovingRangeOffset = Sheet.selectionRange.EndPos;

                        Sheet.RequestInvalidate();
                        isProcessed = true;
                    }

                    #endregion // Hit Selection Drag
                }

                if (!isProcessed
                    && Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToMoveCells))
                {
                    #region Hit Selection Move

                    var selBounds = Sheet.GetRangePhysicsBounds(Sheet.selectionRange);
                    selBounds.Width--;
                    selBounds.Height--;

                    if (GraphicsToolkit.PointOnRectangleBounds(selBounds, location, 2 / ScaleFactor))
                    {
                        Sheet.draggingSelectionRange = Sheet.selectionRange;
                        Sheet.operationStatus = OperationStatus.SelectionRangeMovePrepare;

                        // set offset position (from selection left-top corner to mouse current location)
                        var pos = GetPosByPoint(this, location);

                        // set offset
                        Sheet.lastMouseMoving.Y = Sheet.selectionRange.Row;
                        Sheet.lastMouseMoving.X = Sheet.selectionRange.Col;

                        // make hover position inside selection range
                        if (pos.Row < Sheet.selectionRange.Row) pos.Row = Sheet.selectionRange.Row;
                        if (pos.Col < Sheet.selectionRange.Col) pos.Col = Sheet.selectionRange.Col;
                        if (pos.Row > Sheet.selectionRange.EndRow) pos.Row = Sheet.selectionRange.EndRow;
                        if (pos.Col > Sheet.selectionRange.EndCol) pos.Col = Sheet.selectionRange.EndCol;

                        // set offset
                        Sheet.focusMovingRangeOffset.Row = pos.Row - Sheet.selectionRange.Row;
                        Sheet.focusMovingRangeOffset.Col = pos.Col - Sheet.selectionRange.Col;

                        Sheet.RequestInvalidate();
                        SetFocus();
                        isProcessed = true;
                    }

                    #endregion // Hit Selection Move
                }
            }

            #region Hit Print Breaks

#if PRINT
			// process page break lines adjusting
			if (!isProcessed
				&& sheet.HasSettings(
				// when the page breaks are showing
				WorksheetSettings.View_ShowPageBreaks |
				// when the user inserting or adjusting the page breaks is allowed
				WorksheetSettings.Behavior_AllowUserChangingPageBreaks)
				//&& !sheet.HasSettings(WorksheetSettings.Edit_Readonly)
				)
			{
				int splitCol = sheet.FindBreakIndexOfColumnByPixel(location);
				if (splitCol >= 0)
				{
					sheet.pageBreakAdjustCol = splitCol;
					sheet.pageBreakAdjustFocusIndex = sheet.pageBreakCols[splitCol];
					sheet.lastMouseMoving.X = sheet.pageBreakAdjustFocusIndex;

					sheet.operationStatus = OperationStatus.AdjustPageBreakColumn;
					sheet.RequestInvalidate();
					this.SetFocus();
					isProcessed = true;
				}

				if (!isProcessed)
				{
					int splitRow = sheet.FindBreakIndexOfRowByPixel(location);
					if (splitRow >= 0)
					{
						sheet.pageBreakAdjustRow = splitRow;
						sheet.pageBreakAdjustFocusIndex = sheet.pageBreakRows[splitRow];
						sheet.lastMouseMoving.Y = sheet.pageBreakAdjustFocusIndex;

						sheet.operationStatus = OperationStatus.AdjustPageBreakRow;
						sheet.RequestInvalidate();
						this.SetFocus();
						isProcessed = true;
					}
				}
			}
#endif // PRINT

            #endregion // Hit Print Breaks

            #region Hit Cells

            if (!isProcessed)
            {
                var row = GetRowByPoint(this, location.Y);
#if DEBUG
                Debug.Assert(row >= 0 && row < Sheet.rows.Count);
#endif // DEBUG

                if (row != -1) // in valid rows
                {
                    var col = GetColByPoint(this, location.X);
#if DEBUG
                    Debug.Assert(col >= 0 && col < Sheet.cols.Count);
#endif // DEBUG

                    if (col != -1) // in valid cols
                    {
                        var pos = new CellPosition(row, col);

                        var cell = Sheet.cells[row, col];

                        if (cell != null || Sheet.HasCellMouseDown)
                        {
                            if (cell != null && !cell.IsValidCell) cell = Sheet.GetMergedCellOfRange(cell);

                            if ((cell != null && cell.body != null) || Sheet.HasCellMouseDown)
                            {
                                var cellRect = Sheet.GetCellBounds(pos);

                                var evtArg = new CellMouseEventArgs(Sheet, cell, pos, new Point(
                                    location.X - cellRect.Left,
                                    location.Y - cellRect.Top), location, buttons, 1);

                                Sheet.RaiseCellMouseDown(evtArg);

                                if (cell != null && cell.body != null)
                                    if (cell.body.OnMouseDown(evtArg))
                                    {
                                        isProcessed = true;

                                        // if cell body has processed any mouse down event,
                                        // it is necessary to cancel double click event to Control instance.
                                        //
                                        // this flag use to notify Control to ignore the double click event once.
                                        Sheet.IgnoreMouseDoubleClick = true;

                                        Sheet.RequestInvalidate();

                                        if (cell.body.AutoCaptureMouse() || evtArg.Capture)
                                        {
                                            SetFocus();
                                            Sheet.mouseCapturedCell = cell;
                                            Sheet.operationStatus = OperationStatus.CellBodyCapture;
                                        }
                                    }
                            }
                        }

#if EX_SCRIPT
						object scriptReturn = sheet.RaiseScriptEvent("onmousedown", RSUtility.CreatePosObject(pos));
						if (scriptReturn != null && !ScriptRunningMachine.GetBoolValue(scriptReturn))
						{
							return true;
						}
#endif // EX_SCRIPT

                        if (!isProcessed)
                        {
                            SetFocus();

                            #region Range Selection

                            //else
                            // do not change focus cell if selection mode is null
                            if (Sheet.selectionMode != WorksheetSelectionMode.None)
                                if (
                                    // mouse left button to start new selection session
                                    buttons == MouseButtons.Left
                                    // or mouse right button to show context-menu, that starts also new selection session
                                    || !Sheet.selectionRange.Contains(row, col))
                                {
                                    // if mouse left button pressed, change operation status to free range selection
                                    switch (Sheet.selectionMode)
                                    {
                                        case WorksheetSelectionMode.Row:
                                            Sheet.operationStatus = OperationStatus.FullRowSelect;
                                            break;

                                        case WorksheetSelectionMode.SingleRow:
                                            Sheet.operationStatus = OperationStatus.FullSingleRowSelect;
                                            break;

                                        case WorksheetSelectionMode.Column:
                                            Sheet.operationStatus = OperationStatus.FullColumnSelect;
                                            break;

                                        case WorksheetSelectionMode.SingleColumn:
                                            Sheet.operationStatus = OperationStatus.FullSingleColumnSelect;
                                            break;

                                        case WorksheetSelectionMode.Range:
                                            Sheet.operationStatus = OperationStatus.RangeSelect;
                                            break;
                                    }

                                    Sheet.SelectRangeStartByMouse(PointToController(location));
                                }

                            #endregion // Cell Selection

                            if (buttons == MouseButtons.Right)
                                Sheet.controlAdapter.ShowContextMenuStrip(ViewTypes.None, PointToController(location));

                            // block other processes
                            isProcessed = true;
                        }
                    }
                }
            }

            #endregion // Hit Cells

            return isProcessed;
        }

        public override bool OnMouseMove(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.CellBodyCapture:

                    #region Cell Body Capture

                    if (Sheet.mouseCapturedCell != null && Sheet.mouseCapturedCell.body != null)
                    {
                        var rowTop = Sheet.rows[Sheet.mouseCapturedCell.InternalRow].Top;
                        var colLeft = Sheet.cols[Sheet.mouseCapturedCell.InternalCol].Left;

                        var evtArg = new CellMouseEventArgs(Sheet, Sheet.mouseCapturedCell, new Point(
                            location.X - colLeft,
                            location.Y - rowTop), location, buttons, 1);

                        isProcessed = Sheet.mouseCapturedCell.body.OnMouseMove(evtArg);
                    }

                    #endregion // Cell Body Capture

                    break;

                case OperationStatus.Default:

                    #region Default Cells Hover

                {
                    var cursorChanged = false;

                    if (Sheet.selectionMode != WorksheetSelectionMode.None
                        && !Sheet.HasSettings(WorksheetSettings.Edit_Readonly))
                    {
                        #region Hover - Check to drag serial

                        if (Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToFillSerial))
                            if (SelectDragCornerHitTest(Sheet, location))
                            {
                                Sheet.controlAdapter.ChangeCursor(CursorStyle.Cross);
                                cursorChanged = true;
                                isProcessed = true;
                            }

                        #endregion // Hover - Check to drag selection

                        #region Hover - Check to move selection

                        if (!isProcessed
                            && Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToMoveCells))
                        {
                            var selBounds = Sheet.GetRangePhysicsBounds(Sheet.SelectionRange);
                            selBounds.Width--;
                            selBounds.Height--;

                            var selHover = GraphicsToolkit.PointOnRectangleBounds(selBounds, location, 2 / ScaleFactor);

                            if (selHover)
                            {
                                Sheet.controlAdapter.ChangeCursor(CursorStyle.Move);
                                cursorChanged = true;
                                isProcessed = true;
                            }
                        }

                        #endregion // Hover - Selection range

                        #region Hover - Highlight ranges

                        // process highlight range hover
                        if (!isProcessed
                            && Sheet.highlightRanges != null)
                            foreach (var refRange in Sheet.highlightRanges)
                            {
                                var scaledRangeRect = Sheet.GetRangePhysicsBounds(refRange);
                                scaledRangeRect.Width--;
                                scaledRangeRect.Height--;

                                var hover = GraphicsToolkit.PointOnRectangleBounds(scaledRangeRect, location, 2f);

                                if (hover)
                                {
                                    Sheet.controlAdapter.ChangeCursor(CursorStyle.Move);
                                    cursorChanged = true;
                                }

                                if (hover != refRange.Hover)
                                {
                                    refRange.Hover = hover;
                                    isProcessed = true;
                                    Sheet.RequestInvalidate();
                                    break;
                                }
                            }

                        #endregion // Hover - Highlight ranges
                    }

                    #region Cell Hover - Page breaks

#if PRINT
						// process page break lines hover
						if (!isProcessed
							&& sheet.HasSettings(
							// when the page breaks are showing
							WorksheetSettings.View_ShowPageBreaks |
							// when the user inserting or adjusting page breaks is allowed
							WorksheetSettings.Behavior_AllowUserChangingPageBreaks)
							//&& !sheet.HasSettings(WorksheetSettings.Edit_Readonly)
							)
						{
							int splitCol = sheet.FindBreakIndexOfColumnByPixel(location);
							if (splitCol >= 0)
							{
								sheet.controlAdapter.ChangeCursor(CursorStyle.ResizeHorizontal);
								cursorChanged = true;
								isProcessed = true;
							}

							int splitRow = sheet.FindBreakIndexOfRowByPixel(location);
							if (splitRow >= 0)
							{
								sheet.controlAdapter.ChangeCursor(CursorStyle.ResizeVertical);
								cursorChanged = true;
								isProcessed = true;
							}
						}
#endif // PRINT

                    #endregion // Hover - Page breaks

                    #region Cell Hover - Cells

                    // process cells hover
                    if (!isProcessed)
                    {
                        var newHoverPos = GetPosByPoint(this, location);
                        if (newHoverPos != Sheet.hoverPos) Sheet.HoverPos = newHoverPos;

                        if (!Sheet.hoverPos.IsEmpty)
                        {
                            var cell = Sheet.cells[Sheet.hoverPos.Row, Sheet.hoverPos.Col];

                            if (cell != null || Sheet.HasCellMouseMove)
                            {
                                if (cell != null && !cell.IsValidCell) cell = Sheet.GetMergedCellOfRange(cell);

                                if ((cell != null && cell.body != null) || Sheet.HasCellMouseMove)
                                {
                                    var cellRect = Sheet.GetCellBounds(Sheet.hoverPos);

                                    var evtArg = new CellMouseEventArgs(Sheet, cell, Sheet.hoverPos, new Point(
                                        location.X - cellRect.Left,
                                        location.Y - cellRect.Top), location, buttons, 1);

                                    Sheet.RaiseCellMouseMove(evtArg);

                                    if (cell != null && cell.body != null) cell.body.OnMouseMove(evtArg);

                                    if (evtArg.CursorStyle != CursorStyle.PlatformDefault)
                                    {
                                        cursorChanged = true;
                                        Sheet.controlAdapter.ChangeCursor(evtArg.CursorStyle);
                                    }
                                }
                            }
                        }

                        if (!cursorChanged) Sheet.controlAdapter.ChangeCursor(CursorStyle.Selection);
                    }

                    #endregion // Cell Hover - Cells
                }

                    #endregion // Default Cells Hover

                    break;

                case OperationStatus.AdjustColumnWidth:
                case OperationStatus.AdjustRowHeight:
                    // do nothing
                    break;

                case OperationStatus.SelectionRangeMovePrepare:

                    #region Ready to move selection

                    // prepare to move selection
                    if (buttons == MouseButtons.Left
                        && Sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToMoveCells)
                        && Sheet.draggingSelectionRange != RangePosition.Empty)
                    {
                        Sheet.operationStatus = OperationStatus.SelectionRangeMove;
                        isProcessed = true;
                    }

                    #endregion // Ready to move selection

                    break;

                case OperationStatus.SelectionRangeMove:

                    #region Selection Range Move

                {
                    var pos = GetPosByPoint(this, location);

                    // reuse lastMoseMoving (compare Point to ReoGridPos)
                    if (Sheet.lastMouseMoving.Y != pos.Row || Sheet.lastMouseMoving.X != pos.Col)
                    {
                        Sheet.lastMouseMoving.Y = pos.Row;
                        Sheet.lastMouseMoving.X = pos.Col;

                        Sheet.draggingSelectionRange = new RangePosition(pos.Row - Sheet.focusMovingRangeOffset.Row,
                            pos.Col - Sheet.focusMovingRangeOffset.Col,
                            Sheet.selectionRange.Rows, Sheet.selectionRange.Cols);

                        if (Sheet.draggingSelectionRange.Row < 0) Sheet.draggingSelectionRange.Row = 0;
                        if (Sheet.draggingSelectionRange.Col < 0) Sheet.draggingSelectionRange.Col = 0;

                        // keep range inside spreadsheet - row
                        if (Sheet.draggingSelectionRange.EndRow >= Sheet.RowCount)
                            Sheet.draggingSelectionRange.Row = Sheet.RowCount - Sheet.draggingSelectionRange.Rows;

                        // keep range inside spreadsheet - col
                        if (Sheet.draggingSelectionRange.EndCol >= Sheet.ColumnCount)
                            Sheet.draggingSelectionRange.Col = Sheet.ColumnCount - Sheet.draggingSelectionRange.Cols;

                        Sheet.ScrollToCell(pos);
                        Sheet.RequestInvalidate();
                    }

                    isProcessed = true;
                }

                    #endregion // Selection Range Move

                    break;

                case OperationStatus.DragSelectionFillSerial:

                    #region Selection Range Drag

                {
                    var pos = GetPosByPoint(this, location);

                    // reuse lastMoseMoving (compare Point to ReoGridPos)
                    if (Sheet.focusMovingRangeOffset != pos)
                    {
                        Sheet.focusMovingRangeOffset = pos;

                        var minRow = Math.Min(Sheet.selectionRange.Row, pos.Row);
                        var minCol = Math.Min(Sheet.selectionRange.Col, pos.Col);
                        var maxRow = Math.Max(Sheet.selectionRange.EndRow, pos.Row);
                        var maxCol = Math.Max(Sheet.selectionRange.EndCol, pos.Col);

                        var selLoc = Sheet.GetRangePhysicsBounds(Sheet.selectionRange);

                        var horizontal = true;

                        if (location.X <= selLoc.X)
                        {
                            if (location.Y < selLoc.Y)
                                // left top
                                horizontal = Math.Abs(selLoc.X - location.X) > Math.Abs(selLoc.Y - location.Y);
                            else if (location.Y > selLoc.Bottom)
                                // left bottom
                                horizontal = Math.Abs(selLoc.X - location.X) > Math.Abs(selLoc.Bottom - location.Y);
                        }
                        else if (location.X >= selLoc.Right)
                        {
                            if (location.Y < selLoc.Y)
                                // right top
                                horizontal = Math.Abs(selLoc.Right - location.X) > Math.Abs(selLoc.Y - location.Y);
                            else if (location.Y > selLoc.Bottom)
                                // right bottom
                                horizontal = Math.Abs(selLoc.Right - location.X) > Math.Abs(selLoc.Bottom - location.Y);
                        }
                        else
                        {
                            horizontal = false;
                        }

                        if (horizontal)
                        {
                            minRow = Sheet.selectionRange.Row;
                            maxRow = Sheet.selectionRange.EndRow;
                        }
                        else
                        {
                            minCol = Sheet.selectionRange.Col;
                            maxCol = Sheet.selectionRange.EndCol;
                        }

                        Sheet.draggingSelectionRange = Sheet.FixRange(new RangePosition(minRow, minCol,
                            maxRow - minRow + 1, maxCol - minCol + 1));
                        Sheet.ScrollToCell(pos);

                        Sheet.RequestInvalidate();
                    }

                    isProcessed = true;
                }

                    #endregion // Selection Range Drag

                    break;

#if PRINT
				case OperationStatus.AdjustPageBreakRow:
                #region Page Break Row
					if (buttons == MouseButtons.Left
						&& sheet.pageBreakAdjustRow > -1)
					{
						int rowIndex = sheet.FindRowIndexMiddle(location.Y);
						int index = sheet.FixPageBreakRowIndex(sheet.pageBreakAdjustRow, rowIndex);

						if (sheet.lastMouseMoving.Y != index)
						{
							sheet.pageBreakAdjustFocusIndex = index;
							sheet.lastMouseMoving.Y = index;

							sheet.RequestInvalidate();
						}

						isProcessed = true;
					}
                #endregion // Page Break Row
					break;

				case OperationStatus.AdjustPageBreakColumn:
                #region Page Break Column
					if (buttons == MouseButtons.Left
						&& sheet.pageBreakAdjustCol > -1)
					{
						int colIndex = sheet.FindColIndexMiddle(location.X);
						int index = sheet.FixPageBreakColIndex(sheet.pageBreakAdjustCol, colIndex);

						if (sheet.lastMouseMoving.X != index)
						{
							sheet.pageBreakAdjustFocusIndex = index;
							sheet.lastMouseMoving.X = index;

							sheet.RequestInvalidate();
						}

						isProcessed = true;
					}
                #endregion // Page Break Column
					break;
#endif // PRINT

                case OperationStatus.RangeSelect:
                case OperationStatus.FullRowSelect:
                case OperationStatus.FullColumnSelect:

                    #region Range Select

                    if (buttons == MouseButtons.Left)
                    {
                        var sp = PointToController(location);
                        Sheet.SelectRangeEndByMouse(sp);
                    }

                    #endregion // Range Select

                    break;

                default:
                    Sheet.controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
                    break;
            }

            return isProcessed;
        }

        public override bool OnMouseUp(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.CellBodyCapture:

                    #region CellBodyCapture

                    if (Sheet.mouseCapturedCell != null)
                    {
                        if (Sheet.mouseCapturedCell.body != null)
                        {
                            var rowTop = Sheet.rows[Sheet.mouseCapturedCell.InternalRow].Top;
                            var colLeft = Sheet.cols[Sheet.mouseCapturedCell.InternalCol].Left;

                            var evtArg = new CellMouseEventArgs(Sheet, Sheet.mouseCapturedCell, new Point(
                                location.X - colLeft,
                                location.Y - rowTop), location, buttons, 1);

                            isProcessed = Sheet.mouseCapturedCell.body.OnMouseUp(evtArg);

                            if (isProcessed) Sheet.RequestInvalidate();
                        }

                        Sheet.mouseCapturedCell = null;
                    }

                    Sheet.operationStatus = OperationStatus.Default;

                    #endregion // CellBodyCapture

                    break;

                case OperationStatus.SelectionRangeMovePrepare:

                    #region Abort Selection Range Move

                    Sheet.operationStatus = OperationStatus.Default;
                    Sheet.RequestInvalidate();
                    isProcessed = true;

                    #endregion // Abort Selection Range Move

                    break;

                case OperationStatus.SelectionRangeMove:

                    #region Submit Selection Range Move

                    if (Sheet.selectionRange != Sheet.draggingSelectionRange)
                    {
                        var fromRange = Sheet.selectionRange;
                        var toRange = Sheet.draggingSelectionRange;

                        try
                        {
                            if (Sheet.CheckRangeReadonly(fromRange))
                                throw new RangeContainsReadonlyCellsException(fromRange);

                            if (Sheet.CheckRangeReadonly(toRange))
                                throw new RangeContainsReadonlyCellsException(toRange);

                            BaseWorksheetAction action;

                            if (PlatformUtility.IsKeyDown(KeyCode.ControlKey))
                                action = new CopyRangeAction(fromRange, toRange.StartPos);
                            else
                                action = new MoveRangeAction(fromRange, toRange.StartPos);

                            Sheet.DoAction(action);
                        }
                        catch (Exception ex)
                        {
                            Sheet.NotifyExceptionHappen(ex);
                        }
                    }

                    Sheet.focusMovingRangeOffset = CellPosition.Empty;
                    Sheet.draggingSelectionRange = RangePosition.Empty;
                    Sheet.operationStatus = OperationStatus.Default;
                    Sheet.controlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
                    Sheet.RequestInvalidate();
                    isProcessed = true;

                    #endregion // Submit Selection Range Move

                    break;

#if FORMULA
				case OperationStatus.DragSelectionFillSerial:
                #region Submit Selection Drag
					sheet.operationStatus = OperationStatus.Default;

					if (sheet.draggingSelectionRange.Rows > sheet.selectionRange.Rows
						|| sheet.draggingSelectionRange.Cols > sheet.selectionRange.Cols)
					{
						RangePosition targetRange = RangePosition.Empty;

						if (sheet.draggingSelectionRange.Rows == sheet.selectionRange.Rows)
						{
							targetRange = new RangePosition(
								sheet.draggingSelectionRange.Row,
								sheet.draggingSelectionRange.Col + sheet.selectionRange.Cols,
								sheet.draggingSelectionRange.Rows,
								sheet.draggingSelectionRange.Cols - sheet.selectionRange.Cols);
						}
						else if (sheet.draggingSelectionRange.Cols == sheet.selectionRange.Cols)
						{
							targetRange = new RangePosition(
								sheet.draggingSelectionRange.Row + sheet.selectionRange.Rows,
								sheet.draggingSelectionRange.Col,
								sheet.draggingSelectionRange.Rows - sheet.selectionRange.Rows,
								sheet.draggingSelectionRange.Cols);
						}

						if (targetRange != RangePosition.Empty)
						{
							sheet.DoAction(new AutoFillSerialAction(sheet.SelectionRange, targetRange));
						}
					}

					sheet.RequestInvalidate();
					isProcessed = true;
                #endregion // Submit Selection Drag
					break;
#endif // FORMULA

                case OperationStatus.RangeSelect:
                case OperationStatus.FullRowSelect:
                case OperationStatus.FullColumnSelect:

                    #region Change Selection Range

                {
                    var pos = GetPosByPoint(this, location);

                    if (Sheet.lastChangedSelectionRange != Sheet.selectionRange)
                    {
                        Sheet.lastChangedSelectionRange = Sheet.selectionRange;
                        Sheet.selEnd = pos;

#if WINFORM || WPF
                            //if (sheet.controlAdapter.ControlInstance is IRangePickableControl)
                            //{
                            if (sheet.whenRangePicked != null)
                            {
                                if (sheet.whenRangePicked(sheet, sheet.selectionRange))
                                {
                                    sheet.EndPickRange();
                                }
                            }
                            //}
#endif // WINFORM || WPF

                        Sheet.RaiseSelectionRangeChanged(new RangeEventArgs(Sheet.selectionRange));

#if EX_SCRIPT
							object scriptReturn =
 sheet.RaiseScriptEvent("onmouseup", RSUtility.CreatePosObject(sheet.selEnd));

							// run if script return true or nothing
							if (scriptReturn == null || ScriptRunningMachine.GetBoolValue(scriptReturn))
							{
								sheet.RaiseScriptEvent("onselectionchange");
							}
#endif // EX_SCRIPT
                    }

                    {
                        var row = pos.Row;
                        var col = pos.Col;

                        var cell = Sheet.cells[row, col];

                        if ((cell != null && cell.body != null) || Sheet.HasCellMouseUp)
                        {
                            var rowTop = Sheet.rows[row].Top;
                            var colLeft = Sheet.cols[col].Left;

                            var evtArg = new CellMouseEventArgs(Sheet, cell, pos, new Point(
                                location.X - colLeft,
                                location.Y - rowTop), location, buttons, 1);

                            Sheet.RaiseCellMouseUp(evtArg);

                            if (cell != null && cell.body != null) isProcessed = cell.body.OnMouseUp(evtArg);
                        }
                    }
                    if (Sheet.selectionRange.IsSingleCell)
                        Sheet.StartEdit();

                    Sheet.operationStatus = OperationStatus.Default;
                    isProcessed = true;
                }

                    #endregion // Change Selection Range

                    break;

#if PRINT
				case OperationStatus.AdjustPageBreakColumn:
                #region Adjust Page Break Column
					if (sheet.pageBreakAdjustCol > -1)
					{
						//SetPageBreakColIndex(this.pageBreakAdjustCol, this.commonMouseMoveColIndex);

						int oldIndex = sheet.pageBreakCols[sheet.pageBreakAdjustCol];

						if (oldIndex >= 0)
						{
							if (oldIndex != sheet.pageBreakAdjustFocusIndex)
							{
								sheet.ChangeColumnPageBreak(oldIndex, sheet.pageBreakAdjustFocusIndex);
							}
							else
							{
								sheet.RequestInvalidate();
							}
						}
					}
					sheet.pageBreakAdjustCol = -1;
					sheet.pageBreakAdjustFocusIndex = -1;
					sheet.operationStatus = OperationStatus.Default;
					isProcessed = true;
                #endregion // Adjust Page Break Column
					break;

				case OperationStatus.AdjustPageBreakRow:
                #region Adjust Page Break Row
					if (sheet.pageBreakAdjustRow > -1)
					{
						int oldIndex = sheet.pageBreakRows[sheet.pageBreakAdjustRow];

						if (oldIndex >= 0)
						{
							if (oldIndex != sheet.pageBreakAdjustFocusIndex)
							{
								sheet.ChangeRowPageBreak(oldIndex, sheet.pageBreakAdjustFocusIndex);
							}
							else
							{
								sheet.RequestInvalidate();
							}
						}
					}
					sheet.pageBreakAdjustRow = -1;
					sheet.pageBreakAdjustFocusIndex = -1;
					sheet.operationStatus = OperationStatus.Default;
					isProcessed = true;
                #endregion // Adjust Page Break Row
					break;
#endif // PRINT

                default:

                    #region Call Event CellMouseUp

                {
                    var pos = GetPosByPoint(this, location);

                    var row = pos.Row;
                    var col = pos.Col;

                    var cell = Sheet.cells[row, col];

                    if ((cell != null && cell.body != null) || Sheet.HasCellMouseUp)
                    {
                        var rowTop = Sheet.rows[row].Top;
                        var colLeft = Sheet.cols[col].Left;

                        var evtArg = new CellMouseEventArgs(Sheet, cell, pos, new Point(
                            location.X - colLeft,
                            location.Y - rowTop), location, buttons, 1);

                        Sheet.RaiseCellMouseUp(evtArg);

                        if (cell != null && cell.body != null) isProcessed = cell.body.OnMouseUp(evtArg);
                    }
                }

                    #endregion // Call Event CellMouseUp

                    break;
            }

            return isProcessed;
        }

        #region DoubleClick

        public override bool OnMouseDoubleClick(Point location, MouseButtons buttons)
        {
            if (!Sheet.focusPos.IsEmpty)
            {
                var pos = GetPosByPoint(this, location);

                if (!pos.IsEmpty
                    && pos.Row < Sheet.rows.Count
                    && pos.Col < Sheet.cols.Count)
                {
                    var cell = Sheet.cells[pos.Row, pos.Col];

                    if (cell != null && !cell.IsValidCell) pos = cell.MergeStartPos;
                }

                if (Sheet.focusPos == pos)
                {
                    Sheet.StartEdit();

                    Sheet.controlAdapter.EditControlApplySystemMouseDown();

                    return true;
                }
            }

            return false;
        }

        #endregion // DoubleClick

        #endregion // Mouse

        #region Utility

        internal static int GetColByPoint(IViewport view, double x)
        {
            var sheet = view.ViewportController.Worksheet;

            if (sheet.cols.Count <= 0 || x < sheet.cols[0].Right) return 0;

            var visibleRegion = view.VisibleRegion;

            // view only contain one column
            if (visibleRegion.EndCol <= visibleRegion.StartCol) return visibleRegion.StartCol;

            // binary search to find the column which contains the give position
            return ArrayHelper.QuickFind((visibleRegion.EndCol - visibleRegion.StartCol + 1) / 2,
                0, sheet.cols.Count - 1, i =>
                {
                    var colHeader = sheet.cols[i];

                    if (colHeader.Right < x)
                        return 1;
                    if (colHeader.Left > x)
                        return -1;
                    return 0;
                });
        }

        internal static int GetRowByPoint(IViewport view, double y)
        {
            var sheet = view.ViewportController.Worksheet;

            if (sheet.rows.Count <= 0 || y < sheet.rows[0].Bottom) return 0;

            var visibleRegion = view.VisibleRegion;

            // view only contain one row
            if (visibleRegion.EndRow <= visibleRegion.StartRow) return visibleRegion.StartRow;

#if DEBUG
            var sw = Stopwatch.StartNew();
            try
            {
#endif
                // binary search to find the row which contains the give position
                return ArrayHelper.QuickFind((visibleRegion.EndRow - visibleRegion.StartRow + 1) / 2,
                    0, sheet.rows.Count - 1, i =>
                    {
                        var rowHeader = sheet.rows[i];

                        if (rowHeader.Bottom < y)
                            return 1;
                        if (rowHeader.Top > y)
                            return -1;
                        return 0;
                    });

#if DEBUG
            }
            finally
            {
                sw.Stop();
                var ms = sw.ElapsedMilliseconds;
                if (ms > 1) Debug.WriteLine("finding row index takes " + ms + " ms.");
            }
#endif
        }

        public static CellPosition GetPosByPoint(IViewport view, Point p)
        {
            return new CellPosition(GetRowByPoint(view, p.Y), GetColByPoint(view, p.X));
        }

        /// <summary>
        ///     Transform position of specified cell into the position on control
        /// </summary>
        /// <param name="view">Source view of the specified cell position.</param>
        /// <param name="pos">Cell position to be converted.</param>
        /// <param name="p">Output point of the cell position related to grid control.</param>
        /// <returns>True if conversion is successful; Otherwise return false.</returns>
        public static bool TryGetCellPositionToControl(IView view, CellPosition pos, out Point p)
        {
            if (view == null)
            {
                p = new Point();
                return false;
            }

            var sheet = view.ViewportController.Worksheet;

            if (sheet == null)
            {
                p = new Point();
                return false;
            }

            pos = sheet.FixPos(pos);

            var viewport = view as IViewport;

            if (viewport == null)
                p = new Point(sheet.cols[pos.Col].Left * view.ScaleFactor + view.Left,
                    sheet.rows[pos.Row].Top * view.ScaleFactor + view.Top);
            else
                p = new Point(
                    sheet.cols[pos.Col].Left * view.ScaleFactor + viewport.Left -
                    viewport.ScrollViewLeft * view.ScaleFactor,
                    sheet.rows[pos.Row].Top * view.ScaleFactor + viewport.Top -
                    viewport.ScrollViewTop * view.ScaleFactor);

            return true;
        }

        internal static bool SelectDragCornerHitTest(Worksheet sheet, Point location)
        {
            var selBounds = sheet.GetRangePhysicsBounds(sheet.SelectionRange);
            selBounds.Width--;
            selBounds.Height--;

            var selectionBorderWidth = sheet.controlAdapter.ControlStyle.SelectionBorderWidth;

            var thumbRect = new Rectangle(selBounds.Right - selectionBorderWidth,
                selBounds.Bottom - selectionBorderWidth,
                selectionBorderWidth + 2, selectionBorderWidth + 2);

            return thumbRect.Contains(location);
        }

        public override string ToString()
        {
            return string.Format("CellsViewport[{0}]", ViewBounds);
        }

        #endregion // Utility
    }
}