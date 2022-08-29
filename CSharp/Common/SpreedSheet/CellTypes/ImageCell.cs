using SpreedSheet.Core.Enum;
using SpreedSheet.Rendering;
using System;
using RGImage = System.Windows.Media.ImageSource;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Representation for an image of cell body
    /// </summary>
    public class ImageCell : CellBody
    {
        /// <summary>
        ///     Get or set the image to be displayed in cell
        /// </summary>
        public RGImage Image { get; set; }

        #region OnPaint

        /// <summary>
        ///     Render the image cell body.
        /// </summary>
        /// <param name="dc">Platform no-associated drawing context instance.</param>
        public override void OnPaint(CellDrawingContext dc)
        {
            if (Image != null)
            {
                double x = Bounds.X;
                double y = Bounds.Y;
                double width = 0;
                double height = 0;
                var needClip = false;

                switch (viewMode)
                {
                    default:
                    case ImageCellViewMode.Stretch:
                        width = Bounds.Width;
                        height = Bounds.Height;
                        break;

                    case ImageCellViewMode.Zoom:
                        var widthRatio = Bounds.Width / Image.Width;
                        var heightRatio = Bounds.Height / Image.Height;
                        var minRatio = Math.Min(widthRatio, heightRatio);
                        width = minRatio * Image.Width;
                        height = minRatio * Image.Height;
                        break;

                    case ImageCellViewMode.Clip:
                        width = Image.Width;
                        height = Image.Height;

                        if (width > Bounds.Width || height > Bounds.Height) needClip = true;
                        break;
                }

                switch (Cell.Style.HAlign)
                {
                    default:
                    case GridHorAlign.Left:
                        x = Bounds.X;
                        break;

                    case GridHorAlign.Center:
                        x = (Bounds.Width - width) / 2;
                        break;

                    case GridHorAlign.Right:
                        x = Bounds.Width - width;
                        break;
                }

                switch (Cell.Style.VAlign)
                {
                    default:
                    case GridVerAlign.Top:
                        y = Bounds.Y;
                        break;

                    case GridVerAlign.Middle:
                        y = (Bounds.Height - height) / 2;
                        break;

                    case GridVerAlign.Bottom:
                        y = Bounds.Height - height;
                        break;
                }

                var g = dc.Graphics;

                if (needClip) g.PushClip(Bounds);

                g.DrawImage(Image, x, y, width, height);

                if (needClip) g.PopClip();
            }

            dc.DrawCellText();
        }

        #endregion // OnPaint

        public override ICellBody Clone()
        {
            return new ImageCell(Image);
        }

        #region Constructor

        /// <summary>
        ///     Create image cell object.
        /// </summary>
        public ImageCell()
        {
        }

        /// <summary>
        ///     Construct image cell-body to show a specified image
        /// </summary>
        /// <param name="image">Image to be displayed</param>
        public ImageCell(RGImage image)
            : this(image, default(ImageCellViewMode))
        {
        }

        /// <summary>
        ///     Construct image cell-body to show a image by specified display-method
        /// </summary>
        /// <param name="image">Image to be displayed</param>
        /// <param name="viewMode">View mode decides how to display a image inside a cell</param>
        public ImageCell(RGImage image, ImageCellViewMode viewMode)
        {
            Image = image;
            this.viewMode = viewMode;
        }

        #endregion // Constructor

        #region ViewMode

        protected ImageCellViewMode viewMode;

        /// <summary>
        ///     Set or get the view mode of this image cell
        /// </summary>
        public ImageCellViewMode ViewMode
        {
            get { return viewMode; }
            set
            {
                if (viewMode != value)
                {
                    viewMode = value;

                    if (Cell != null && Cell.Worksheet != null) Cell.Worksheet.RequestInvalidate();
                }
            }
        }

        #endregion // ViewMode
    }

    #region ImageCellViewMode

    /// <summary>
    ///     Image dispaly method in ImageCell-body
    /// </summary>
    public enum ImageCellViewMode
    {
        /// <summary>
        ///     Fill to cell boundary. (default)
        /// </summary>
        Stretch,

        /// <summary>
        ///     Lock aspect ratio to fit cell boundary.
        /// </summary>
        Zoom,

        /// <summary>
        ///     Keep original image size and clip to fill the cell.
        /// </summary>
        Clip
    }

    #endregion // ImageCellViewMode
}