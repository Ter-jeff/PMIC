using System;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.CellTypes
{
    #region NegativeProgressCell

    /// <summary>
    ///     Progress bar for display both positive and negative percent.
    /// </summary>
    [Serializable]
    public class NegativeProgressCell : CellBody
    {
        #region Attributes

        /// <summary>
        ///     Get or set color for positive display.
        /// </summary>
        public SolidColor PositiveColor { get; set; }

        /// <summary>
        ///     Get or set color for negative display.
        /// </summary>
        public SolidColor NegativeColor { get; set; }

        /// <summary>
        ///     Determines whether or not display a linear gradient color on progress bar.
        /// </summary>
        public bool LinearGradient { get; set; }

        /// <summary>
        ///     Determines whether or not display the cell text or value.
        /// </summary>
        public bool DisplayCellText { get; set; }

        /// <summary>
        ///     Determines whether or not force to display the progress inside cell.
        /// </summary>
        public bool LimitedInsideCell { get; set; }

        #endregion // Attributes

        #region Constructor

        /// <summary>
        ///     Create negative progress cell.
        /// </summary>
        public NegativeProgressCell()
        {
            PositiveColor = SolidColor.LightGreen;
            NegativeColor = SolidColor.LightCoral;
            LinearGradient = true;
            DisplayCellText = true;
            LimitedInsideCell = true;
        }

        #endregion // Constructor

        #region OnPaint

        /// <summary>
        ///     Render the negative progress cell body.
        /// </summary>
        /// <param name="dc"></param>
        public override void OnPaint(CellDrawingContext dc)
        {
            var value = Cell.GetData<double>();

            if (LimitedInsideCell)
            {
                if (value > 1) value = 1;
                else if (value < -1) value = -1;
            }

            var g = dc.Graphics;

            Rectangle rect;

            if (value >= 0)
            {
                rect = new Rectangle(Bounds.Left + Bounds.Width / 2, Bounds.Top + 1,
                    Bounds.Width * (value / 2.0d), Bounds.Height - 1);

                if (rect.Width > 0 && rect.Height > 0)
                {
                    if (LinearGradient)
                        g.FillRectangleLinear(PositiveColor,
                            new SolidColor(0, PositiveColor), 0, rect);
                    else
                        g.FillRectangle(rect, PositiveColor);
                }
            }
            else
            {
                var center = Bounds.Left + Bounds.Width / 2.0f;
                var left = Bounds.Width * value * 0.5d;
                rect = new Rectangle(center + left, Bounds.Top + 1, -left, Bounds.Height - 1);

                if (rect.Width > 0 && rect.Height > 0)
                {
                    if (LinearGradient)
                        g.FillRectangleLinear(NegativeColor,
                            new SolidColor(0, NegativeColor), 180, rect);
                    else
                        g.FillRectangle(rect, NegativeColor);
                }
            }

            if (DisplayCellText) dc.DrawCellText();
        }

        #endregion // OnPaint

        #endregion // NegativeProgressCell
    }
}