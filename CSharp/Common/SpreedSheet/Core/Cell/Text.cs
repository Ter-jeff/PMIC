#define WPF


#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WPF

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
    }

    partial class Cell
    {
        private Rectangle textBounds;
        internal bool FontDirty { get; set; }

        /// <summary>
        ///     text boundary for display
        /// </summary>
        internal Rectangle TextBounds
        {
            get { return textBounds; }
            set { textBounds = value; }
        }

        internal double TextBoundsTop
        {
            get { return textBounds.Y; }
            set { textBounds.Y = value; }
        }

        internal double TextBoundsLeft
        {
            get { return textBounds.X; }
            set { textBounds.X = value; }
        }

        internal double TextBoundsWidth
        {
            get { return textBounds.Width; }
            set { textBounds.Width = value; }
        }

        internal double TextBoundsHeight
        {
            get { return textBounds.Height; }
            set { textBounds.Height = value; }
        }

        internal Rectangle PrintTextBounds { get; set; }

        /// <summary>
        ///     Horizontal alignement for display
        /// </summary>
        internal GridRenderHorAlign RenderHorAlign { get; set; }

        /// <summary>
        ///     Column span if text larger than the cell it inside
        /// </summary>
        internal short RenderTextColumnSpan { get; set; }

        //private SolidColor renderColor = null;
        /// <summary>
        ///     Get the render color of cell text. Render color is the final color that used to render the text on the worksheet.
        ///     Whatever cell style with text color is specified, negative numbers may displayed as red.
        ///     This property cannot be changed directly.
        ///     To change text color, set cell style with text color by call SetRangeStyles method, or change the
        ///     Cell.Style.TextColor property.
        /// </summary>
        public SolidColor RenderColor { get; internal set; }

        //public SolidColor? GetRenderColor() { return renderColor; }

        internal double RenderScaleFactor { get; set; }
    }
}