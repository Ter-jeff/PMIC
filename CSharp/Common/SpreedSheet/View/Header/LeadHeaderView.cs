using SpreedSheet.Core;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Header
{
    internal class LeadHeaderView : View
    {
        protected Worksheet Sheet;

        public LeadHeaderView(ViewportController vc)
            : base(vc)
        {
            Sheet = vc.Worksheet;
        }

        #region Draw

        public override void Draw(CellDrawingContext dc)
        {
            if (Bounds.Width <= 0 || Bounds.Height <= 0 || Sheet.controlAdapter == null) return;

            var g = dc.Graphics;
            var controlStyle = Sheet.workbook.controlAdapter.ControlStyle;

            g.FillRectangle(Bounds, controlStyle.Colors[ControlAppearanceColors.LeadHeadNormal]);

            var startColor = Sheet.isLeadHeadSelected
                ? controlStyle.Colors[ControlAppearanceColors.LeadHeadIndicatorStart]
                : controlStyle.Colors[ControlAppearanceColors.LeadHeadSelected];

            var endColor = controlStyle.Colors[ControlAppearanceColors.LeadHeadIndicatorEnd];

            dc.Renderer.DrawLeadHeadArrow(Bounds, startColor, endColor);
        }

        #endregion // Draw

        public override bool OnMouseDown(Point location, MouseButtons buttons)
        {
            // mouse down in LeadHead?
            switch (Sheet.operationStatus)
            {
                case OperationStatus.Default:
                    if (Sheet.selectionMode != WorksheetSelectionMode.None)
                    {
                        Sheet.SelectRange(RangePosition.EntireRange);

                        // show context menu
                        if (buttons == MouseButtons.Right)
                            Sheet.controlAdapter.ShowContextMenuStrip(ViewTypes.LeadHeader, location);

                        return true;
                    }

                    break;
            }

            return false;
        }

        public override bool OnMouseMove(Point location, MouseButtons buttons)
        {
            Sheet.controlAdapter.ChangeCursor(CursorStyle.Selection);

            return false;
        }
    }
}