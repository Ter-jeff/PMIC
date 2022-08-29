using System;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Header
{
    internal class ColumnHeaderView : HeaderView
    {
        public ColumnHeaderView(IViewportController vc)
            : base(vc)
        {
            ScrollableDirections = ScrollDirection.Horizontal;
        }

        #region Draw

        public override void DrawView(CellDrawingContext dc)
        {
            var r = dc.Renderer;
            var g = dc.Graphics;

            if (Bounds.Height <= 0 || Sheet.controlAdapter == null) return;

            var controlStyle = Sheet.workbook.controlAdapter.ControlStyle;

            r.BeginDrawHeaderText(ScaleFactor);

            var splitterLinePen = r.GetPen(controlStyle.Colors[ControlAppearanceColors.RowHeadSplitter]);
            var headerTextBrush = r.GetBrush(controlStyle.Colors[ControlAppearanceColors.ColHeadText]);

            var isFullColSelected = Sheet.SelectionRange.Rows == Sheet.RowCount;

            for (var i = VisibleRegion.StartCol; i <= VisibleRegion.EndCol; i++)
            {
                var isSelected = i >= Sheet.SelectionRange.Col && i <= Sheet.SelectionRange.EndCol;

                var header = Sheet.cols[i];

                var x = header.Left * ScaleFactor;
                var width = header.InnerWidth * ScaleFactor;

                if (!header.IsVisible)
                {
                    g.DrawLine(splitterLinePen, x - 1, 0, x - 1, Bounds.Bottom);
                }
                else
                {
                    var rect = new Rectangle(x, 0, width, Bounds.Height);

#if WINFORM || WPF
                    g.FillRectangleLinear(
                        controlStyle.GetColHeadStartColor(false, isSelected, isSelected && isFullColSelected, false),
                        controlStyle.GetColHeadEndColor(false, isSelected, isSelected && isFullColSelected, false), 90f,
                        rect);
#elif ANDROID
					g.FillRectangle(rect, controlStyle.GetRowHeadEndColor(false, isSelected, isSelected && isFullColSelected, false));
#endif // ANDROID

                    g.DrawLine(splitterLinePen, x, 0, x, Bounds.Height);

                    var textBrush = header.TextColor != null
                        ? dc.Renderer.GetBrush((SolidColor)header.TextColor)
                        : headerTextBrush;

                    if (textBrush == null) textBrush = headerTextBrush;

                    r.DrawHeaderText(header.RenderText, textBrush, rect);

                    if (header.Body != null)
                    {
                        g.PushTransform();
                        g.TranslateTransform(rect.X, rect.Y);
                        header.Body.OnPaint(dc, rect.Size);
                        g.PopTransform();
                    }
                }
            }

            var lx = Sheet.cols[VisibleRegion.EndCol].Right * ScaleFactor;
            g.DrawLine(splitterLinePen, lx, 0, lx, Bounds.Height);

            //g.DrawLine(splitterLinePen, this.ViewLeft, Bounds.Height, this.ViewLeft + Bounds.Width, Bounds.Height);

            // bottom line
            //if (!sheet.HasSettings(WorksheetSettings.View_ShowGuideLine))
            //{
            //	g.DrawLine(ViewLeft, Bounds.Bottom,
            //		Math.Min((sheet.cols[sheet.cols.Count - 1].Right - ViewLeft) * this.ScaleFactor + Bounds.Left, Bounds.Width),
            //		//ViewLeft+ Bounds.Width,
            //		Bounds.Bottom, controlStyle.Colors[ControlAppearanceColors.ColHeadSplitter]);
            //}
        }

        #endregion

        #region Utility

        public static Rectangle GetColHeaderBounds(Worksheet sheet, int col, Point position)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");

            var viewportController = sheet.ViewportController;

            if (viewportController == null || viewportController.View == null)
                throw new ArgumentNullException("viewportController");

            //viewportController.Bounds

            IViewport view = viewportController.View.GetViewByPoint(position) as ColumnHeaderView;

            if (view == null)
                throw new ArgumentNullException("Cannot found column header view from specified position");

            if (view is ColumnHeaderView)
            {
                var header = sheet.RetrieveColumnHeader(col);

                var ScaleFactor = sheet.renderScaleFactor;

                return new Rectangle(header.Left * ScaleFactor + view.Left - view.ScrollViewLeft,
                    view.Top - view.ScrollViewTop,
                    header.InnerWidth * ScaleFactor,
                    sheet.ColHeaderHeight * ScaleFactor);
            }

            return new Rectangle();
        }

        #endregion

        #region Mouse

        public override bool OnMouseDown(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            switch (Sheet.operationStatus)
            {
                case OperationStatus.Default:
                    var col = -1;
                    var inSeparator = Sheet.FindColumnByPosition(location.X, out col);

                    if (col >= 0)
                    {
                        // adjust columns width
                        if (inSeparator
                            && buttons == MouseButtons.Left
                            && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustColumnWidth))
                        {
                            Sheet.currentColWidthChanging = col;
                            Sheet.operationStatus = OperationStatus.AdjustColumnWidth;
                            Sheet.controlAdapter.ChangeCursor(CursorStyle.ChangeColumnWidth);
                            Sheet.RequestInvalidate();

                            HeaderAdjustBackup = Sheet.headerAdjustNewValue =
                                Sheet.cols[Sheet.currentColWidthChanging].InnerWidth;
                            SetFocus();

                            isProcessed = true;
                        }

                        if (!isProcessed)
                        {
                            var header = Sheet.cols[col];

                            if (header.Body != null)
                            {
                                // let body to decide the mouse behavior
                                var arg = new WorksheetMouseEventArgs(Sheet, new Point(
                                        (location.X - header.Left) * ScaleFactor,
                                        location.Y / ScaleFactor),
                                    new Point((location.X - header.Left) * ScaleFactor + Left,
                                        location.Y / ScaleFactor), buttons, 1);

                                isProcessed = header.Body.OnMouseDown(
                                    new Size(header.InnerWidth * ScaleFactor, Sheet.ColHeaderHeight), arg);
                            }
                        }

                        if (!isProcessed
                            // do not allow to select column if selection mode is null
                            && Sheet.selectionMode != WorksheetSelectionMode.None)
                        {
                            var isFullColSelected =
                                Sheet.selectionMode == WorksheetSelectionMode.Range
                                && Sheet.selectionRange.Rows == Sheet.rows.Count
                                && Sheet.selectionRange.ContainsColumn(col);

                            // select whole column
                            if (!isFullColSelected || buttons == MouseButtons.Left)
                            {
                                Sheet.operationStatus = OperationStatus.FullColumnSelect;
                                Sheet.controlAdapter.ChangeCursor(CursorStyle.FullColumnSelect);

                                SetFocus();

                                Sheet.SelectRangeStartByMouse(PointToController(location));

                                isProcessed = true;
                            }
                        }

                        // show context menu
                        if (buttons == MouseButtons.Right)
                            Sheet.ControlAdapter.ShowContextMenuStrip(ViewTypes.ColumnHeader,
                                PointToController(location));
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
                case OperationStatus.AdjustColumnWidth:
                    if (Sheet.currentColWidthChanging >= 0
                        && buttons == MouseButtons.Left)
                    {
                        var colHeader = Sheet.cols[Sheet.currentColWidthChanging];
                        Sheet.headerAdjustNewValue = location.X - colHeader.Left;
                        if (Sheet.headerAdjustNewValue < 0) Sheet.headerAdjustNewValue = 0;

                        Sheet.controlAdapter.ChangeCursor(CursorStyle.ChangeColumnWidth);
                        Sheet.RequestInvalidate();
                        isProcessed = true;
                    }

                    break;

                case OperationStatus.Default:
                {
                    if (Sheet.currentColWidthChanging == -1 && Sheet.currentRowHeightChanging == -1)
                    {
                        var col = -1;

                        // find the column index
                        var inline = Sheet.FindColumnByPosition(location.X, out col)
                                     && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustColumnWidth);

                        if (col >= 0)
                        {
                            var curStyle = inline ? CursorStyle.ChangeColumnWidth :
                                Sheet.selectionMode == WorksheetSelectionMode.None ? CursorStyle.Selection :
                                CursorStyle.FullColumnSelect;

                            var header = Sheet.cols[col];

                            // check if header body exists
                            if (header.Body != null)
                            {
                                // let cell's body decide the mouse behavior
                                var arg = new WorksheetMouseEventArgs(Sheet, new Point(
                                    (location.X - header.Left) * ScaleFactor,
                                    location.Y / ScaleFactor), location, buttons, 1)
                                {
                                    CursorStyle = curStyle
                                };

                                isProcessed = header.Body.OnMouseMove(
                                    new Size(header.InnerWidth * ScaleFactor, Sheet.ColHeaderHeight), arg);

                                curStyle = arg.CursorStyle;
                            }

                            Sheet.controlAdapter.ChangeCursor(curStyle);
                        }
                    }
                }
                    break;

                case OperationStatus.FullColumnSelect:
                case OperationStatus.FullSingleColumnSelect:
                    if (buttons == MouseButtons.Left)
                    {
                        Sheet.controlAdapter.ChangeCursor(CursorStyle.FullColumnSelect);
                        Sheet.SelectRangeEndByMouse(PointToController(location));

                        isProcessed = true;
                    }

                    break;
            }

            return isProcessed;
        }

        public override bool OnMouseUp(Point location, MouseButtons buttons)
        {
            switch (Sheet.operationStatus)
            {
                case OperationStatus.AdjustColumnWidth:
                    if (Sheet.currentColWidthChanging > -1)
                    {
                        SetColumnsWidthAction setColsWidthAction;

                        var isFullColSelected = Sheet.selectionMode == WorksheetSelectionMode.Range
                                                && Sheet.selectionRange.Rows == Sheet.rows.Count
                                                && Sheet.selectionRange.ContainsColumn(Sheet.currentColWidthChanging);

                        var targetWidth = (ushort)Sheet.headerAdjustNewValue;

                        if (targetWidth != HeaderAdjustBackup)
                        {
                            if (isFullColSelected)
                                setColsWidthAction = new SetColumnsWidthAction(Sheet.selectionRange.Col,
                                    Sheet.selectionRange.Cols, targetWidth);
                            else
                                setColsWidthAction =
                                    new SetColumnsWidthAction(Sheet.currentColWidthChanging, 1, targetWidth);

                            Sheet.DoAction(setColsWidthAction);
                        }
                    }

                    Sheet.currentColWidthChanging = -1;
                    Sheet.operationStatus = OperationStatus.Default;
                    Sheet.RequestInvalidate();

                    HeaderAdjustBackup = Sheet.headerAdjustNewValue = 0;
                    FreeFocus();

                    return true;

                case OperationStatus.FullColumnSelect:
                    Sheet.operationStatus = OperationStatus.Default;
                    Sheet.ControlAdapter.ChangeCursor(CursorStyle.Selection);
                    FreeFocus();
                    return true;
            }

            return false;
        }

        public override bool OnMouseDoubleClick(Point location, MouseButtons buttons)
        {
            var col = -1;
            var inSeparator = Sheet.FindColumnByPosition(location.X, out col);
            if (col >= 0)
                // adjust columns width
                if (inSeparator
                    && buttons == MouseButtons.Left
                    && Sheet.HasSettings(WorksheetSettings.Edit_AllowAdjustColumnWidth))
                {
                    Sheet.AutoFitColumnWidth(col, true);

                    return true;
                }

            return false;
        }

        #endregion
    }
}