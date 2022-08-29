using System.Collections.Generic;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using Point = unvell.ReoGrid.Graphics.Point;
using Size = unvell.ReoGrid.Graphics.Size;

namespace SpreedSheet.Rendering
{
    internal class WpfGraphics : IGraphics
    {
        protected ResourcePoolManager ResourceManager = new ResourcePoolManager();

        public ResourcePoolManager ResourcePoolManager
        {
            get { return ResourceManager; }
        }

        public System.Windows.Media.DrawingContext PlatformGraphics { get; set; }

        public void DrawImage(ImageSource image, double x, double y, double width, double height)
        {
            PlatformGraphics.DrawImage(image, new Rect(x, y, width, height));
        }

        public void DrawImage(ImageSource image, Rectangle bounds)
        {
            PlatformGraphics.DrawImage(image, bounds);
        }

        public void FillEllipse(IColor fillColor, Rectangle rect)
        {
            var b = ResourceManager.GetBrush(fillColor.ToSolidColor());

            if (b != null) PlatformGraphics.DrawEllipse(b, null, rect.Origin, rect.Width, rect.Height);
        }

        public void FillPolygon(Point[] points, SolidColor startColor, SolidColor endColor, double angle,
            Rectangle rect)
        {
            var pts = new System.Windows.Point[points.Length];
            for (var i = 0; i < pts.Length; i++) pts[i] = points[i];

            var streamGeometry = new StreamGeometry();
            using (var geometryContext = streamGeometry.Open())
            {
                geometryContext.PolyLineTo(pts, true, true);
            }

            var lgb = new LinearGradientBrush(startColor, endColor, angle);
            PlatformGraphics.DrawGeometry(lgb, null, streamGeometry);
        }

        internal DashStyle ToWpfDashStyle(LineStyles style)
        {
            switch (style)
            {
                default:
                case LineStyles.Solid: return DashStyles.Solid;
                case LineStyles.Dot: return DashStyles.Dot;
                case LineStyles.Dash: return DashStyles.Dash;
                case LineStyles.DashDot: return DashStyles.DashDot;
                case LineStyles.DashDotDot: return DashStyles.DashDotDot;
            }
        }

        #region Line

        public void DrawLine(Pen p, double x1, double y1, double x2, double y2)
        {
            DrawLine(p, new Point(x1, y1), new Point(x2, y2));
        }

        public void DrawLine(Pen p, Point startPoint, Point endPoint)
        {
            var halfPenWidth = p.Thickness / 2;

            // Create a guidelines set
            var guidelines = new GuidelineSet();

            guidelines.GuidelinesX.Add(startPoint.X + halfPenWidth);
            guidelines.GuidelinesY.Add(startPoint.Y + halfPenWidth);

            PlatformGraphics.PushGuidelineSet(guidelines);

            PlatformGraphics.DrawLine(p, startPoint, endPoint);

            PlatformGraphics.Pop();
        }

        public void DrawLine(Point startPoint, Point endPoint, SolidColor color)
        {
            PlatformGraphics.DrawLine(ResourceManager.GetPen(color), startPoint, endPoint);
        }

        public void DrawLine(double x1, double y1, double x2, double y2, SolidColor color)
        {
            var pen = ResourceManager.GetPen(color);
            DrawLine(pen, x1, y1, x2, y2);
        }

        public void DrawLine(double x1, double y1, double x2, double y2, SolidColor color, double width,
            LineStyles style)
        {
            var p = ResourceManager.GetPen(color, width, ToWpfDashStyle(style));

            if (p != null) PlatformGraphics.DrawLine(p, new System.Windows.Point(x1, y1), new System.Windows.Point(x2, y2));
        }

        public void DrawLine(Point startPoint, Point endPoint, SolidColor color, double width, LineStyles style)
        {
            var p = ResourceManager.GetPen(color, width, ToWpfDashStyle(style));

            if (p != null) PlatformGraphics.DrawLine(p, startPoint, endPoint);
        }

        public void DrawLines(Point[] points, int start, int length, SolidColor color, double width, LineStyles style)
        {
            if (!color.IsTransparent && length > 1)
            {
                var p = ResourceManager.GetPen(color, width, ToWpfDashStyle(style));

                if (p != null)
                {
                    var geo = new PathGeometry();
                    for (int i = 1, k = start + 1; i < length; i++, k++)
                        geo.AddGeometry(new LineGeometry(points[k - 1], points[k]));
                    PlatformGraphics.DrawGeometry(null, p, geo);
                }
            }
        }

        #endregion // Line

        #region Rectangle

        public void DrawRectangle(Pen p, Rectangle rect)
        {
            PlatformGraphics.DrawRectangle(null, p, rect);
        }

        public void DrawRectangle(Pen p, double x, double y, double w, double h)
        {
            PlatformGraphics.DrawRectangle(null, p, new Rect(x, y, w, h));
        }

        public void DrawRectangle(Rectangle rect, SolidColor color)
        {
            var p = ResourceManager.GetPen(color);
            if (p != null) PlatformGraphics.DrawRectangle(null, p, rect);
        }

        public void DrawRectangle(double x, double y, double width, double height, SolidColor color)
        {
            var p = ResourceManager.GetPen(color);
            PlatformGraphics.DrawRectangle(null, p, new Rect(x, y, width, height));
        }

        public void FillRectangle(HatchStyles style, SolidColor hatchColor, SolidColor bgColor, Rectangle rect)
        {
            // TODO
        }

        public void FillRectangle(HatchStyles style, SolidColor hatchColor, SolidColor bgColor, double x, double y,
            double width, double height)
        {
            // TODO
        }

        public void FillRectangle(Rectangle rect, IColor color)
        {
            if (color is SolidColor)
                PlatformGraphics.DrawRectangle(ResourceManager.GetBrush((SolidColor)color), null, rect);
        }

        public void FillRectangle(double x, double y, double width, double height, IColor color)
        {
            if (color is SolidColor)
                PlatformGraphics.DrawRectangle(ResourceManager.GetBrush((SolidColor)color), null,
                    new Rect(x, y, width, height));
        }

        public void FillRectangle(Brush b, double x, double y, double width, double height)
        {
            PlatformGraphics.DrawRectangle(b, null, new Rect(x, y, width, height));
        }

        public void FillRectangleLinear(SolidColor color1, SolidColor color2, double angle, Rectangle rect)
        {
            var lgb = new LinearGradientBrush(color1, color2, angle);
            PlatformGraphics.DrawRectangle(lgb, null, rect);
        }


        public void DrawRectangle(Rectangle rect, SolidColor color, double width, LineStyles lineStyle)
        {
            var p = ResourceManager.GetPen(color, width, ToWpfDashStyle(lineStyle));
            if (p != null) PlatformGraphics.DrawRectangle(null, p, rect);
        }

        public void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor)
        {
            if (fillColor is SolidColor)
                PlatformGraphics.DrawRectangle(ResourceManager.GetBrush((SolidColor)fillColor),
                    ResourceManager.GetPen(lineColor), rect);
        }

        public void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor, double width,
            LineStyles lineStyle)
        {
            var p = ResourceManager.GetPen(lineColor, width, ToWpfDashStyle(lineStyle));
            var b = ResourceManager.GetBrush(fillColor.ToSolidColor());

            if (p != null && b != null) PlatformGraphics.DrawRectangle(b, p, rect);
        }

        #endregion // Rectangle

        #region Text

        public void DrawText(string text, string fontName, double size, SolidColor color, Rectangle rect)
        {
            DrawText(text, fontName, size, color, rect, ReoGridHorAlign.Left, GridVerAlign.Top);
        }

        public void DrawText(string text, string fontName, double size, SolidColor color, Rectangle rect,
            ReoGridHorAlign halign, GridVerAlign valign)
        {
            if (rect.Width > 0 && rect.Height > 0 && !string.IsNullOrEmpty(text))
            {
                var ft = new FormattedText(text, Thread.CurrentThread.CurrentCulture,
                    FlowDirection.LeftToRight, ResourceManager.GetTypeface(fontName),
                    size * PlatformUtility.GetDPI() / 72.0,
                    ResourceManager.GetBrush(color));

                ft.MaxTextWidth = rect.Width;
                ft.MaxTextHeight = rect.Height;

                switch (halign)
                {
                    case ReoGridHorAlign.Left:
                        ft.TextAlignment = TextAlignment.Left;
                        break;

                    case ReoGridHorAlign.Center:
                        ft.TextAlignment = TextAlignment.Center;
                        break;

                    case ReoGridHorAlign.Right:
                        ft.TextAlignment = TextAlignment.Right;
                        break;
                }

                switch (valign)
                {
                    case GridVerAlign.Middle:
                        rect.Y += (rect.Height - ft.Height) / 2;
                        break;

                    case GridVerAlign.Bottom:
                        rect.Y += rect.Height - ft.Height;
                        break;
                }

                PlatformGraphics.DrawText(ft, rect.Location);
            }
        }

        public Size MeasureText(string text, string fontName, double fontSize, Size displayArea)
        {
            // in WPF environment do not measure text, use FormattedText instead
            return new Size(0, 0);
        }

        #endregion // Text

        #region Clip

        public void PushClip(Rectangle clipRect)
        {
            PlatformGraphics.PushClip(new RectangleGeometry(clipRect));
        }

        public void PopClip()
        {
            PlatformGraphics.Pop();
        }

        #endregion // Clip

        #region Transform

        private readonly Stack<MatrixTransform> _transformStack = new Stack<MatrixTransform>();

        public void PushTransform()
        {
            PushTransform(Matrix.Identity);
        }

        public void PushTransform(Matrix m)
        {
            var mt = new MatrixTransform(m);
            _transformStack.Push(mt);
            PlatformGraphics.PushTransform(mt);
        }

        Matrix IGraphics.PopTransform()
        {
            PlatformGraphics.Pop();
            return _transformStack.Pop().Matrix;
        }

        public Matrix PopTransform()
        {
            PlatformGraphics.Pop();
            return _transformStack.Pop().Matrix;
        }

        public void TranslateTransform(double x, double y)
        {
            if (_transformStack.Count > 0)
            {
                var mt = _transformStack.Peek();
                var m2 = new Matrix();
                m2.Translate(x, y);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void ScaleTransform(double x, double y)
        {
            if (x != 0 && y != 0
                       && x != 1 && y != 1
                       && _transformStack.Count > 0)
            {
                var mt = _transformStack.Peek();
                var m2 = new Matrix();
                m2.Scale(x, y);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void RotateTransform(double angle)
        {
            if (_transformStack.Count > 0)
            {
                var mt = _transformStack.Peek();
                var m = mt.Matrix;
                var m2 = new Matrix();
                m2.RotateAt(angle, m.OffsetX, m.OffsetY);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void ResetTransform()
        {
            if (_transformStack.Count > 0)
            {
                var mt = _transformStack.Peek();
                mt.Matrix = Matrix.Identity;
            }
        }

        #endregion // Transform

        #region Ellipse

        public void DrawEllipse(SolidColor color, Rectangle rectangle)
        {
            var p = ResourceManager.GetPen(color);
            if (p != null)
                PlatformGraphics.DrawEllipse(null, p, new Point(rectangle.X + rectangle.Width / 2,
                    rectangle.Y + rectangle.Height / 2), rectangle.Width, rectangle.Height);
        }

        public void DrawEllipse(SolidColor color, double x, double y, double width, double height)
        {
            var p = ResourceManager.GetPen(color);
            if (p != null) PlatformGraphics.DrawEllipse(null, p, new Point(x, y), width, height);
        }

        public void DrawEllipse(Pen pen, Rectangle rectangle)
        {
            PlatformGraphics.DrawEllipse(null, pen, rectangle.Location, rectangle.Width, rectangle.Height);
        }

        public void FillEllipse(Brush b, Rectangle rectangle)
        {
            PlatformGraphics.DrawEllipse(b, null, rectangle.Location, rectangle.Width, rectangle.Height);
        }

        public void FillEllipse(Brush b, double x, double y, double width, double height)
        {
            PlatformGraphics.DrawEllipse(b, null, new Point(x, y), width, height);
        }

        #endregion // Ellipse

        #region Polygon

        public void DrawPolygon(SolidColor color, double width, LineStyles style, params Point[] points)
        {
            DrawLines(points, 0, points.Length, color, width, style);
        }

        public void FillPolygon(IColor color, params Point[] points)
        {
            if (!color.IsTransparent)
            {
                var geo = new PathGeometry();

                for (int i = 1, k = 1; i < points.Length; i++, k++)
                    geo.AddGeometry(new LineGeometry(points[k - 1], points[k]));

                PlatformGraphics.DrawGeometry(new SolidColorBrush(color.ToSolidColor()), null, geo);
            }
        }

        #endregion // Polygon

        #region Utility

        public bool IsAntialias
        {
            get { return true; }
            set { }
        }

        public void Reset()
        {
            _transformStack.Clear();
        }

        internal void SetPlatformGraphics(System.Windows.Media.DrawingContext dc)
        {
            PlatformGraphics = dc;
        }

        #endregion // Utility

        #region Path

        public void FillPath(IColor color, Geometry graphicsPath)
        {
            var b = ResourceManager.GetBrush(color.ToSolidColor());
            if (b != null) PlatformGraphics.DrawGeometry(b, null, graphicsPath);
        }

        public void DrawPath(SolidColor color, Geometry graphicsPath)
        {
            var p = ResourceManager.GetPen(color);
            if (p != null) PlatformGraphics.DrawGeometry(null, p, graphicsPath);
        }

        #endregion // Path
    }
}