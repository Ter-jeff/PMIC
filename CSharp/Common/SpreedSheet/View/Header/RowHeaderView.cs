using System;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Header
{
    internal class RowHeaderView : HeaderView
    {
        public RowHeaderView(IViewportController vc)
            : base(vc)
        {
            ScrollableDirections = ScrollDirection.Vertical;
        }

        public override Point PointToView(Point p)
        {
            return base.PointToView(p);
        }

        #region Draw

        public override void DrawView(CellDrawingContext dc)
        {
            var g = dc.Renderer;

            if (Bounds.Width <= 0 || Sheet.controlAdapter == null) return;

            var controlStyle = Sheet.workbook.controlAdapter.ControlStyle;

            g.BeginDrawHeaderText(ScaleFactor);

            var splitterLinePen = dc.Renderer.GetPen(controlStyle.Colors[ControlAppearanceColors.RowHeadSplitter]);
            var defaultTextBrush = dc.Renderer.GetBrush(controlStyle.Colors[ControlAppearanceColors.RowHeadText]);

            var isFullRowSelected = Sheet.SelectionRange.Cols == Sheet.ColumnCount;

            for (var i = VisibleRegion.StartRow; i <= VisibleRegion.EndRow; i++)
            {
                var isSelected = i >= Sheet.SelectionRange.Row && i <= Sheet.SelectionRange.EndRow;

                var row = Sheet.rows[i];
                var y = row.Top * ScaleFactor;

                if (!row.IsVisible)
                {
                    g.DrawLine(splitterLinePen, 0, y - 1, Bounds.Width, y - 1);
                }
                else
                {
                    var rect = new Rectangle(0, y, Bounds.Width, row.InnerHeight * ScaleFactor);

                    if (rect.Height > 0)
                    {
                        g.FillRectangle(rect,
                            controlStyle.GetRowHeadEndColor(false, isSelected, isSelected && isFullRowSelected, false));
                        g.DrawLine(splitterLinePen, new Point(0, y), new Point(Bounds.Width, y));

                        var headerText = row.Text != null ? row.Text : (row.Row + 1).ToString();

                        if (!string.IsNullOrEmpty(headerText))
                        {
                            var textBrush = row.TextColor != null
                                ? dc.Renderer.GetBrush((SolidColor)row.TextColor)
                                : defaultTextBrush;

                            if (textBrush == null) textBrush = defaultTextBrush;

                            g.DrawHeaderText(headerText, textBrush, rect);
                        }

                        if (row.Body != null)
                        {
                            g.PushTransform();
                            g.TranslateTransform(rect.X, rect.Y);
                            row.Body.OnPaint(dc, rect.Size);
                            g.PopTransform();
                        }
                    }
                }
            }

            if (VisibleRegion.EndRow >= 0)
            {
                var ly = Sheet.rows[VisibleRegion.EndRow].Bottom * ScaleFactor;
                g.DrawLine(splitterLinePen, 0, ly, Bounds.Width, ly);
            }

            // right line
            if (!Sheet.HasSettings(WorksheetSettings.View_ShowGridLine))
                dc.Graphics.DrawLine(dc.Renderer.GetPen(controlStyle.Colors[ControlAppearanceColors.RowHeadSplitter]),
                    Bounds.Right, Bounds.Y, Bounds.Right,
                    Math.Min((Sheet.rows[Sheet.rows.Count - 1].Bottom - ScrollViewTop) * ScaleFactor + Bounds.Top,
                        Bounds.Bottom));
        }

        #endregion

        #region Mouse

        public override bool OnMouseDown(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            var row = -1;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.Default:

                    var inSeparator = Sheet.FindRowByPosition(location.Y, out row);

                    if (row >= 0 && row < Sheet.rows.Count)
                    {
                        if (inSeparator
                            && buttons == MouseButtons.Left
                            && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustRowHeight))
                        {
                            Sheet.currentRowHeightChanging = row;
                            Sheet.operationStatus = OperationStatus.AdjustRowHeight;
                            Sheet.controlAdapter.ChangeCursor(CursorStyle.ChangeRowHeight);
                            Sheet.RequestInvalidate();

                            HeaderAdjustBackup = Sheet.headerAdjustNewValue =
                                Sheet.rows[Sheet.currentRowHeightChanging].InnerHeight;
                            SetFocus();

                            isProcessed = true;
                        }
                        else if (Sheet.selectionMode != WorksheetSelectionMode.None)
                        {
                            // check whether entire row is selected, select row if not
                            var isFullRowSelected = Sheet.selectionMode == WorksheetSelectionMode.Range
                                                    && Sheet.selectionRange.Cols == Sheet.cols.Count
                                                    && Sheet.selectionRange.ContainsRow(row);

                            if (!isFullRowSelected || buttons == MouseButtons.Left)
                            {
                                Sheet.operationStatus = OperationStatus.FullRowSelect;
                                Sheet.controlAdapter.ChangeCursor(CursorStyle.FullRowSelect);

                                SetFocus();

                                Sheet.SelectRangeStartByMouse(PointToController(location));

                                isProcessed = true;
                            }

                            if (buttons == MouseButtons.Right)
                                Sheet.controlAdapter.ShowContextMenuStrip(ViewTypes.RowHeader,
                                    PointToController(location));
                        }
                    }

                    break;
            }

            return isProcessed;
        }

        public override bool OnMouseMove(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.Default:
                    if (buttons == MouseButtons.None
                        && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustRowHeight))
                    {
                        var row = -1;
                        var inline = Sheet.FindRowByPosition(location.Y, out row);

                        if (row >= 0)
                            Sheet.controlAdapter.ChangeCursor(inline ? CursorStyle.ChangeRowHeight :
                                Sheet.selectionMode == WorksheetSelectionMode.None ? CursorStyle.PlatformDefault :
                                CursorStyle.FullRowSelect);
                    }

                    break;

                case OperationStatus.AdjustRowHeight:
                    if (buttons == MouseButtons.Left
                        && Sheet.currentRowHeightChanging >= 0)
                    {
                        var rowHeader = Sheet.rows[Sheet.currentRowHeightChanging];
                        Sheet.headerAdjustNewValue = location.Y - rowHeader.Top;
                        if (Sheet.headerAdjustNewValue < 0) Sheet.headerAdjustNewValue = 0;

                        Sheet.controlAdapter.ChangeCursor(CursorStyle.ChangeRowHeight);
                        Sheet.RequestInvalidate();
                        isProcessed = true;
                    }

                    break;

                case OperationStatus.FullRowSelect:
                    Sheet.SelectRangeEndByMouse(PointToController(location));
                    Sheet.controlAdapter.ChangeCursor(CursorStyle.FullRowSelect);

                    isProcessed = true;
                    break;
            }

            return isProcessed;
        }

        public override bool OnMouseUp(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.AdjustRowHeight:
                    if (Sheet.currentRowHeightChanging > -1)
                    {
                        SetRowsHeightAction setRowsHeightAction;

                        var isFullRowSelected = Sheet.selectionMode == WorksheetSelectionMode.Range
                                                && Sheet.selectionRange.Cols == Sheet.cols.Count
                                                && Sheet.selectionRange.ContainsRow(Sheet.currentRowHeightChanging);

                        var targetHeight = (ushort)Sheet.headerAdjustNewValue;

                        if (targetHeight != HeaderAdjustBackup)
                        {
                            if (isFullRowSelected)
                                setRowsHeightAction = new SetRowsHeightAction(Sheet.selectionRange.Row,
                                    Sheet.selectionRange.Rows, targetHeight);
                            else
                                setRowsHeightAction =
                                    new SetRowsHeightAction(Sheet.currentRowHeightChanging, 1, targetHeight);

                            Sheet.DoAction(setRowsHeightAction);
                        }

                        Sheet.currentRowHeightChanging = -1;
                        Sheet.operationStatus = OperationStatus.Default;
                        HeaderAdjustBackup = Sheet.headerAdjustNewValue = 0;

                        Sheet.RequestInvalidate();
                        FreeFocus();
                        isProcessed = true;
                    }

                    break;

                case OperationStatus.FullRowSelect:
                case OperationStatus.FullSingleRowSelect:
                    Sheet.operationStatus = OperationStatus.Default;
                    Sheet.controlAdapter.ChangeCursor(CursorStyle.Selection);

                    FreeFocus();
                    isProcessed = true;
                    break;
            }

            return isProcessed;
        }

        public override bool OnMouseDoubleClick(Point location, MouseButtons buttons)
        {
            var row = -1;
            var inSeparator = Sheet.FindRowByPosition(location.Y, out row);

            if (row >= 0)
                // adjust row height
                if (inSeparator
                    && buttons == MouseButtons.Left
                    && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustRowHeight))
                {
                    Sheet.AutoFitRowHeight(row, true);

                    return true;
                }

            return false;
        }

        #endregion
    }
}