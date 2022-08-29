using RGImage = System.Windows.Media.ImageSource;
using System;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Represents an image button cell on worksheet.
    /// </summary>
    [Serializable]
    public class ImageButtonCell : ButtonCell
    {
        /// <summary>
        ///     Create image button cell without image specified.
        /// </summary>
        public ImageButtonCell()
            : this(null)
        {
        }

        /// <summary>
        ///     Create image button cell with specified image.
        /// </summary>
        /// <param name="image"></param>
        public ImageButtonCell(RGImage image)
        {
            Image = image;
        }

        /// <summary>
        ///     Image that displayed on button.
        /// </summary>
        public RGImage Image { get; set; }

        /// <summary>
        ///     Paint image button cell.
        /// </summary>
        /// <param name="dc">Platform non-associated drawing context.</param>
        public override void OnPaint(CellDrawingContext dc)
        {
            base.OnPaint(dc);

            if (Image != null)
            {
                var widthScale = Math.Min((Bounds.Width - 4) / Image.Width, 1);
                var heightScale = Math.Min((Bounds.Height - 4) / Image.Height, 1);

                var minScale = Math.Min(widthScale, heightScale);
                var imageScale = Image.Height / Image.Width;
                var width = Image.Width * minScale;

                var r = new Rectangle(0, 0, width, imageScale * width);

                r.X = (Bounds.Width - r.Width) / 2;
                r.Y = (Bounds.Height - r.Height) / 2;

                if (IsPressed)
                {
                    r.X++;
                    r.Y++;
                }

                dc.Graphics.DrawImage(Image, r);
            }
        }

        /// <summary>
        ///     Clone image button from this object.
        /// </summary>
        /// <returns>New instance of image button.</returns>
        public override ICellBody Clone()
        {
            return new ImageButtonCell(Image);
        }
    }
}