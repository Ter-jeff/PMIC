using System;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.CellTypes
{
    #region Progress

    /// <summary>
    ///     Representation for a button of cell body
    /// </summary>
    [Serializable]
    public class ProgressCell : CellBody
    {
        /// <summary>
        ///     Create progress cell body.
        /// </summary>
        public ProgressCell()
        {
            TopColor = SolidColor.LightSkyBlue;
            BottomColor = SolidColor.SkyBlue;
        }

        /// <summary>
        ///     Get or set the top color.
        /// </summary>
        public SolidColor TopColor { get; set; }

        /// <summary>
        ///     Get or set the bottom color.
        /// </summary>
        public SolidColor BottomColor { get; set; }

        /// <summary>
        ///     Render the progress cell body.
        /// </summary>
        /// <param name="dc"></param>
        public override void OnPaint(CellDrawingContext dc)
        {
            var value = Cell.GetData<double>();

            if (value > 0)
            {
                var g = dc.Graphics.PlatformGraphics;

                var rect = new Rectangle(Bounds.Left, Bounds.Top + 1, Bounds.Width * value, Bounds.Height - 1);

                if (rect.Width > 0 && rect.Height > 0)
                    dc.Graphics.FillRectangleLinear(TopColor, BottomColor, 90f, rect);
            }
        }

        /// <summary>
        ///     Clone a progress bar from this object.
        /// </summary>
        /// <returns>New instance of progress bar.</returns>
        public override ICellBody Clone()
        {
            return new ProgressCell();
        }
    }

    #endregion // Progress
}