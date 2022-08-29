using SpreedSheet.Core;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Utility;

namespace SpreedSheet.View
{
    internal class CellsForegroundView : Viewport
    {
        public CellsForegroundView(IViewportController vc) : base(vc)
        {
        }

        public override IView GetViewByPoint(Point p)
        {
            // return null to always avoid making this view active
            return null;
        }

        public override void Draw(CellDrawingContext dc)
        {
            var sheet = ViewportController.Worksheet;
            if (sheet == null || sheet.controlAdapter == null) return;

            var g = dc.Graphics;
            var controlStyle = sheet.workbook.controlAdapter.ControlStyle;

            switch (sheet.operationStatus)
            {
                case OperationStatus.AdjustColumnWidth:

                    #region Draw Column Header Adjust Line

                    if (sheet.currentColWidthChanging >= 0)
                    {
                        var col = sheet.cols[sheet.currentColWidthChanging];

                        var left = col.Left * ScaleFactor; // -ViewLeft * this.ScaleFactor;
                        var right = (col.Left + sheet.headerAdjustNewValue) *
                                    ScaleFactor; // -ViewLeft * this.ScaleFactor;
                        var top = ScrollViewTop * ScaleFactor;
                        var bottom = ScrollViewTop * ScaleFactor + Height;

                        g.DrawLine(left, top, left, bottom, SolidColor.Black, 1, LineStyles.Dot);
                        g.DrawLine(right, top, right, bottom, SolidColor.Black, 1, LineStyles.Dot);
                    }

                    #endregion // Draw Column Header Adjust Line

                    break;

                case OperationStatus.AdjustRowHeight:

                    #region Draw Row Header Adjust Line

                    if (sheet.currentRowHeightChanging >= 0)
                    {
                        var row = sheet.rows[sheet.currentRowHeightChanging];

                        var top = row.Top * ScaleFactor;
                        var bottom = (row.Top + sheet.headerAdjustNewValue) * ScaleFactor;
                        var left = ScrollViewLeft * ScaleFactor;
                        var right = ScrollViewLeft * ScaleFactor + Width;

                        g.DrawLine(left, top, right, top, SolidColor.Black, 1, LineStyles.Dot);
                        g.DrawLine(left, bottom, right, bottom, SolidColor.Black, 1, LineStyles.Dot);
                    }

                    #endregion // Draw Row Header Adjust Line

                    break;

                case OperationStatus.DragSelectionFillSerial:
                case OperationStatus.SelectionRangeMovePrepare:
                case OperationStatus.SelectionRangeMove:

                    #region Selection Moving

                    if (sheet.draggingSelectionRange != RangePosition.Empty
                        && dc.DrawMode == DrawMode.View
                        && sheet.HasSettings(WorksheetSettings.Edit_DragSelectionToMoveCells))
                    {
                        var scaledSelectionMovingRect = CellsViewport.GetScaledAndClippedRangeRect(this,
                            sheet.draggingSelectionRange.StartPos,
                            sheet.draggingSelectionRange.EndPos,
                            controlStyle.SelectionBorderWidth);

                        scaledSelectionMovingRect.Offset(-1, -1);

                        var selectionBorderColor = controlStyle.Colors[ControlAppearanceColors.SelectionBorder];

                        dc.Graphics.DrawRectangle(scaledSelectionMovingRect,
                            ColorUtility.FromAlphaColor(255, selectionBorderColor),
                            controlStyle.SelectionBorderWidth, LineStyles.Solid);
                    }

                    #endregion // Selection Moving

                    break;
            }
        }
    }
}