#define WPF


#pragma warning disable 1591

#if WINFORM
using RGFloat = System.Single;

using RGPen = System.Drawing.Pen;
using RGBrush = System.Drawing.Brush;

using RGPath = System.Drawing.Drawing2D.GraphicsPath;
using RGImage = System.Drawing.Image;

using PlatformGraphics = System.Drawing.Graphics;
using RGTransform = System.Drawing.Drawing2D.Matrix;

#elif ANDROID
using RGFloat = System.Single;
using PlatformGraphics = Android.Graphics.Canvas;
using RGPen = Android.Graphics.Paint;
using RGBrush = Android.Graphics.Paint;
using RGPath = Android.Graphics.Path;
using RGImage = Android.Graphics.Picture;
using RGTransform = Android.Graphics.Matrix;

#elif WPF

using SpreedSheet.Core.Enum;
using RGFloat = System.Double;
using RGPath = System.Windows.Media.Geometry;
using RGImage = System.Windows.Media.ImageSource;
using RGPen = System.Windows.Media.Pen;
using RGBrush = System.Windows.Media.Brush;
using PlatformGraphics = System.Windows.Media.DrawingContext;
using RGTransform = System.Windows.Media.Matrix;

#elif iOS
using RGFloat = System.Double;
using PlatformGraphics = CoreGraphics.CGContext;
using RGPen = CoreGraphics.CGContext;
using RGBrush = CoreGraphics.CGContext;
using RGPath = CoreGraphics.CGPath;
using RGImage = CoreGraphics.CGImage;
using RGTransform = CoreGraphics.CGAffineTransform;
#endif // WPF

namespace unvell.ReoGrid.Graphics
{
    /// <summary>
    ///     Represents abstract cross-platform drawing context.
    /// </summary>
    public interface IGraphics
    {
        PlatformGraphics PlatformGraphics { get; set; }

        bool IsAntialias { get; set; }

        void DrawLine(double x1, double y1, double x2, double y2, SolidColor color);
        void DrawLine(double x1, double y1, double x2, double y2, SolidColor color, double width, LineStyles style);
        void DrawLine(Point startPoint, Point endPoint, SolidColor color);

        void DrawLine(Point startPoint, Point endPoint, SolidColor color, double width, LineStyles style);

        //void DrawLine(SolidColor color, Point startPoint, Point endPoint, RGFloat width, LineStyles style, LineCapStyles startCap, LineCapStyles endCap);
        void DrawLine(RGPen p, double x1, double y1, double x2, double y2);
        void DrawLine(RGPen p, Point startPoint, Point endPoint);
        void DrawLines(Point[] points, int start, int length, SolidColor color, double width, LineStyles style);

        void DrawRectangle(Rectangle rect, SolidColor color);
        void DrawRectangle(Rectangle rect, SolidColor color, double width, LineStyles lineStyle);
        void DrawRectangle(double x, double y, double width, double height, SolidColor color);
        void DrawRectangle(RGPen p, Rectangle rect);
        void DrawRectangle(RGPen p, double x, double y, double width, double height);

        void FillRectangle(HatchStyles style, SolidColor hatchColor, SolidColor bgColor, Rectangle rect);

        void FillRectangle(HatchStyles style, SolidColor hatchColor, SolidColor bgColor, double x, double y,
            double width, double height);

        void FillRectangle(Rectangle rect, IColor color);
        void FillRectangle(double x, double y, double width, double height, IColor color);
        void FillRectangle(RGBrush b, double x, double y, double width, double height);
        void FillRectangleLinear(SolidColor startColor, SolidColor endColor, double angle, Rectangle rect);

        void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor);

        void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor, double weight,
            LineStyles lineStyle);

        void DrawEllipse(SolidColor color, Rectangle rectangle);
        void DrawEllipse(SolidColor color, double x, double y, double width, double height);
        void DrawEllipse(RGPen pen, Rectangle rectangle);
        void FillEllipse(IColor fillColor, Rectangle rectangle);
        void FillEllipse(RGBrush b, Rectangle rectangle);
        void FillEllipse(RGBrush b, double x, double y, double widht, double height);

        void DrawPolygon(SolidColor color, double lineWidth, LineStyles lineStyle, params Point[] points);
        void FillPolygon(IColor color, params Point[] points);

        void FillPath(IColor color, RGPath graphicsPath);
        void DrawPath(SolidColor color, RGPath graphicsPath);

        void DrawImage(RGImage image, double x, double y, double width, double height);
        void DrawImage(RGImage image, Rectangle rect);

        void DrawText(string text, string fontName, double size, SolidColor color, Rectangle rect,
            GridHorAlign halign = GridHorAlign.Center, GridVerAlign valign = GridVerAlign.Middle);

        //Graphics.Size MeasureText(string text, string fontName, RGFloat size, Size areaSize);
        //void FillPolygon(RGPointF[] points, RGColor startColor, RGColor endColor, RGFloat angle, RGRectF rect);

        void ScaleTransform(double sx, double sy);
        void TranslateTransform(double x, double y);
        void RotateTransform(double angle);
        void ResetTransform();

        void PushClip(Rectangle clip);
        void PopClip();

        void PushTransform();
        void PushTransform(RGTransform t);
        RGTransform PopTransform();

        void Reset();
    }

    internal struct LineCap
    {
        public LineCapStyles StartStyle { get; set; }
        public LineCapStyles EndStyle { get; set; }

        public Size StartSize { get; set; }
        public Size EndSize { get; set; }

        public SolidColor StartColor { get; set; }
        public SolidColor EndColor { get; set; }
    }
}

#pragma warning restore 1591