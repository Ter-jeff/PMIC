#define WPF

#if WPF
//#define GRID_GUIDELINE

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using SpreedSheet.Core.Enum;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Rendering;
using unvell.Common;
using unvell.ReoGrid.Graphics;
using FontStyles = unvell.ReoGrid.Drawing.Text.FontStyles;
using Point = unvell.ReoGrid.Graphics.Point;
using RGFont = System.Windows.Media.Typeface;
using Size = unvell.ReoGrid.Graphics.Size;
using WPFDrawingContext = System.Windows.Media.DrawingContext;
using WPFPoint = System.Windows.Point;

namespace unvell.ReoGrid.Rendering
{
    #region Graphics

    internal class WPFGraphics : IGraphics
    {
        protected ResourcePoolManager resourceManager = new ResourcePoolManager();

        public ResourcePoolManager ResourcePoolManager
        {
            get { return resourceManager; }
        }

        public WPFDrawingContext PlatformGraphics { get; set; }

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
            var b = resourceManager.GetBrush(fillColor.ToSolidColor());

            if (b != null) PlatformGraphics.DrawEllipse(b, null, rect.Origin, rect.Width, rect.Height);
        }

        public void FillPolygon(Point[] points, SolidColor startColor, SolidColor endColor, double angle,
            Rectangle rect)
        {
            var pts = new WPFPoint[points.Length];
            for (var i = 0; i < pts.Length; i++) pts[i] = points[i];

            var streamGeometry = new StreamGeometry();
            using (var geometryContext = streamGeometry.Open())
            {
                geometryContext.PolyLineTo(pts, true, true);
            }

            var lgb = new LinearGradientBrush(startColor, endColor, angle);
            PlatformGraphics.DrawGeometry(lgb, null, streamGeometry);
        }

        internal DashStyle ToWPFDashStyle(LineStyles style)
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
#if !GRID_GUIDELINE
            var halfPenWidth = p.Thickness / 2;

            // Create a guidelines set
            var guidelines = new GuidelineSet();

            guidelines.GuidelinesX.Add(startPoint.X + halfPenWidth);
            guidelines.GuidelinesY.Add(startPoint.Y + halfPenWidth);

            PlatformGraphics.PushGuidelineSet(guidelines);
#endif // GRID_GUIDELINE

            PlatformGraphics.DrawLine(p, startPoint, endPoint);

#if !GRID_GUIDELINE
            PlatformGraphics.Pop();
#endif // GRID_GUIDELINE
        }

        public void DrawLine(Point startPoint, Point endPoint, SolidColor color)
        {
            PlatformGraphics.DrawLine(resourceManager.GetPen(color), startPoint, endPoint);
        }

        public void DrawLine(double x1, double y1, double x2, double y2, SolidColor color)
        {
            var pen = resourceManager.GetPen(color);
            DrawLine(pen, x1, y1, x2, y2);
        }

        public void DrawLine(double x1, double y1, double x2, double y2, SolidColor color, double width,
            LineStyles style)
        {
            var p = resourceManager.GetPen(color, width, ToWPFDashStyle(style));

            if (p != null) PlatformGraphics.DrawLine(p, new WPFPoint(x1, y1), new WPFPoint(x2, y2));
        }

        public void DrawLine(Point startPoint, Point endPoint, SolidColor color, double width, LineStyles style)
        {
            var p = resourceManager.GetPen(color, width, ToWPFDashStyle(style));

            if (p != null) PlatformGraphics.DrawLine(p, startPoint, endPoint);
        }

        public void DrawLines(Point[] points, int start, int length, SolidColor color, double width, LineStyles style)
        {
            if (!color.IsTransparent && length > 1)
            {
                var p = resourceManager.GetPen(color, width, ToWPFDashStyle(style));

                if (p != null)
                {
                    var geo = new PathGeometry();
                    for (int i = 1, k = start + 1; i < length; i++, k++)
                        geo.AddGeometry(new LineGeometry(points[k - 1], points[k]));
                    PlatformGraphics.DrawGeometry(null, p, geo);
                }
            }
        }

        //public void DrawLine(SolidColor color, Point startPoint, Point endPoint, double width, LineStyles style, LineCapStyles startCap, LineCapStyles endCap)
        //{
        //	var b = this.resourceManager.GetBrush(color);

        //	var p = new Pen(b, width);

        //	if (startCap == LineCapStyles.Arrow)
        //	{
        //		p.StartLineCap = PenLineCap.Triangle;
        //	}

        //	if (endCap == LineCapStyles.Arrow)
        //	{
        //		p.EndLineCap = PenLineCap.Triangle;
        //	}

        //	this.g.DrawLine(p, startPoint, endPoint);
        //}

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
            var p = resourceManager.GetPen(color);
            if (p != null) PlatformGraphics.DrawRectangle(null, p, rect);
        }

        public void DrawRectangle(double x, double y, double width, double height, SolidColor color)
        {
            var p = resourceManager.GetPen(color);
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
                PlatformGraphics.DrawRectangle(resourceManager.GetBrush((SolidColor)color), null, rect);
        }

        public void FillRectangle(double x, double y, double width, double height, IColor color)
        {
            if (color is SolidColor)
                PlatformGraphics.DrawRectangle(resourceManager.GetBrush((SolidColor)color), null,
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
            var p = resourceManager.GetPen(color, width, ToWPFDashStyle(lineStyle));
            if (p != null) PlatformGraphics.DrawRectangle(null, p, rect);
        }

        public void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor)
        {
            if (fillColor is SolidColor)
                PlatformGraphics.DrawRectangle(resourceManager.GetBrush((SolidColor)fillColor),
                    resourceManager.GetPen(lineColor), rect);
        }

        public void DrawAndFillRectangle(Rectangle rect, SolidColor lineColor, IColor fillColor, double width,
            LineStyles lineStyle)
        {
            var p = resourceManager.GetPen(lineColor, width, ToWPFDashStyle(lineStyle));
            var b = resourceManager.GetBrush(fillColor.ToSolidColor());

            if (p != null && b != null) PlatformGraphics.DrawRectangle(b, p, rect);
        }

        #endregion // Rectangle

        #region Text

        public void DrawText(string text, string fontName, double size, SolidColor color, Rectangle rect)
        {
            DrawText(text, fontName, size, color, rect, GridHorAlign.Left, GridVerAlign.Top);
        }

        public void DrawText(string text, string fontName, double size, SolidColor color, Rectangle rect,
            GridHorAlign halign, GridVerAlign valign)
        {
            if (rect.Width > 0 && rect.Height > 0 && !string.IsNullOrEmpty(text))
            {
                var ft = new FormattedText(text, Thread.CurrentThread.CurrentCulture,
                    FlowDirection.LeftToRight, resourceManager.GetTypeface(fontName),
                    size * PlatformUtility.GetDPI() / 72.0,
                    resourceManager.GetBrush(color));

                ft.MaxTextWidth = rect.Width;
                ft.MaxTextHeight = rect.Height;

                switch (halign)
                {
                    case GridHorAlign.Left:
                        ft.TextAlignment = TextAlignment.Left;
                        break;

                    case GridHorAlign.Center:
                        ft.TextAlignment = TextAlignment.Center;
                        break;

                    case GridHorAlign.Right:
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

        private readonly Stack<MatrixTransform> transformStack = new Stack<MatrixTransform>();

        public void PushTransform()
        {
            PushTransform(Matrix.Identity);
        }

        public void PushTransform(Matrix m)
        {
            var mt = new MatrixTransform(m);
            transformStack.Push(mt);
            PlatformGraphics.PushTransform(mt);
        }

        Matrix IGraphics.PopTransform()
        {
            PlatformGraphics.Pop();
            return transformStack.Pop().Matrix;
        }

        public Matrix PopTransform()
        {
            PlatformGraphics.Pop();
            return transformStack.Pop().Matrix;
        }

        public void TranslateTransform(double x, double y)
        {
            if (transformStack.Count > 0)
            {
                var mt = transformStack.Peek();
                var m2 = new Matrix();
                m2.Translate(x, y);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void ScaleTransform(double x, double y)
        {
            if (x != 0 && y != 0
                       && x != 1 && y != 1
                       && transformStack.Count > 0)
            {
                var mt = transformStack.Peek();
                var m2 = new Matrix();
                m2.Scale(x, y);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void RotateTransform(double angle)
        {
            if (transformStack.Count > 0)
            {
                var mt = transformStack.Peek();
                var m = mt.Matrix;
                var m2 = new Matrix();
                m2.RotateAt(angle, m.OffsetX, m.OffsetY);
                mt.Matrix = m2 * mt.Matrix;
            }
        }

        public void ResetTransform()
        {
            if (transformStack.Count > 0)
            {
                var mt = transformStack.Peek();
                mt.Matrix = Matrix.Identity;
            }
        }

        #endregion // Transform

        #region Ellipse

        public void DrawEllipse(SolidColor color, Rectangle rectangle)
        {
            var p = resourceManager.GetPen(color);
            if (p != null)
                PlatformGraphics.DrawEllipse(null, p, new Point(rectangle.X + rectangle.Width / 2,
                    rectangle.Y + rectangle.Height / 2), rectangle.Width, rectangle.Height);
        }

        public void DrawEllipse(SolidColor color, double x, double y, double width, double height)
        {
            var p = resourceManager.GetPen(color);
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
            transformStack.Clear();
        }

        internal void SetPlatformGraphics(WPFDrawingContext dc)
        {
            PlatformGraphics = dc;
        }

        #endregion // Utility

        #region Path

        public void FillPath(IColor color, Geometry graphicsPath)
        {
            var b = resourceManager.GetBrush(color.ToSolidColor());
            if (b != null) PlatformGraphics.DrawGeometry(b, null, graphicsPath);
        }

        public void DrawPath(SolidColor color, Geometry graphicsPath)
        {
            var p = resourceManager.GetPen(color);
            if (p != null) PlatformGraphics.DrawGeometry(null, p, graphicsPath);
        }

        #endregion // Path
    }

    #endregion // Graphics

    #region Renderer

    internal class WPFRenderer : WPFGraphics, IRenderer
    {
        private Pen cachePen;

        private Pen capLinePen;

        private double headerTextScale = 9d;
        protected Typeface headerTextTypeface;

        private LineCap lineCap;

        internal WPFRenderer()
        {
            headerTextTypeface = PlatformUtility.GetFontDefaultTypeface(SystemFonts.SmallCaptionFontFamily);
        }

        public ResourcePoolManager GetResourcePoolManager
        {
            get { return resourceManager; }
        }

        public Size MeasureCellText(Cell cell, DrawMode drawMode, double scale)
        {
            if (cell.InnerStyle.RotationAngle != 0)
            {
                var m = Matrix.Identity;
                double hw = cell.formattedText.Width * 0.5, hh = cell.formattedText.Height * 0.5;
                WPFPoint p1 = new WPFPoint(-hw, -hh), p2 = new WPFPoint(hw, hh);
                m.Rotate(cell.InnerStyle.RotationAngle);
                p1 *= m;
                p2 *= m;
                return new Size(Math.Abs(p1.X - p2.X), Math.Abs(p1.Y - p2.Y));
            }

            return new Size(cell.formattedText.Width, cell.formattedText.Height);
        }

        public void DrawCellText(Cell cell, SolidColor textColor, DrawMode drawMode, double scale)
        {
            var sheet = cell.Worksheet;

            if (sheet == null)
                return;

            if (cell.InnerStyle.RotationAngle != 0)
            {
                var m = Matrix.Identity;
                m.Rotate(cell.InnerStyle.RotationAngle);
                m.Translate(cell.Bounds.OriginX * sheet.ScaleFactor, cell.Bounds.OriginY * sheet.ScaleFactor);
                PushTransform(m);
                PlatformGraphics.DrawText(cell.formattedText,
                    new WPFPoint(-cell.formattedText.Width * 0.5, -cell.formattedText.Height * 0.5));
                PopTransform();
            }
            else
            {
                PlatformGraphics.DrawText(cell.formattedText, cell.TextBounds.Location);
            }
        }

        public void UpdateCellRenderFont(Cell cell, UpdateFontReason reason)
        {
            var sheet = cell.Worksheet;
            if (sheet == null || sheet.controlAdapter == null) return;

            var dpi = PlatformUtility.GetDPI();
            var fontSize = cell.InnerStyle.FontSize * sheet.renderScaleFactor * dpi / 72.0;
            if (cell.formattedText == null || cell.formattedText.Text != cell.InnerDisplay)
            {
                SolidColor textColor = DecideTextColor(cell);
                cell.formattedText = new FormattedText(cell.InnerDisplay,
                    CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    resourceManager.GetTypeface(cell.InnerStyle.FontName),
                    fontSize, resourceManager.GetBrush(textColor));
            }
            else if (reason == UpdateFontReason.FontChanged || reason == UpdateFontReason.ScaleChanged)
            {
                cell.formattedText.SetFontFamily(cell.InnerStyle.FontName);
                cell.formattedText.SetFontSize(fontSize);
            }
            else if (reason == UpdateFontReason.TextColorChanged)
            {
                SolidColor textColor = DecideTextColor(cell);
                cell.formattedText.SetForegroundBrush(resourceManager.GetBrush(textColor));
            }

            var ft = cell.formattedText;


            if (reason == UpdateFontReason.FontChanged || reason == UpdateFontReason.ScaleChanged)
            {
                ft.SetFontWeight(cell.InnerStyle.Bold ? FontWeights.Bold : FontWeights.Normal);

                ft.SetFontStyle(PlatformUtility.ToWPFFontStyle(cell.InnerStyle.fontStyles));

                ft.SetTextDecorations(PlatformUtility.ToWPFFontDecorations(cell.InnerStyle.fontStyles));
            }
        }

        public void DrawRunningFocusRect(double x, double y, double w, double h, SolidColor color, int runningOffset)
        {
        }

        public void BeginCappedLine(LineCapStyles startCap, Size startSize, LineCapStyles endCap, Size endSize,
            SolidColor color, double width)
        {
            capLinePen = new Pen(new SolidColorBrush(color), width);
            capLinePen.StartLineCap = PlatformUtility.ToWPFLineCap(startCap);
            capLinePen.EndLineCap = PlatformUtility.ToWPFLineCap(endCap);
        }

        public void DrawCappedLine(double x1, double y1, double x2, double y2)
        {
            if (capLinePen != null) base.DrawLine(capLinePen, x1, y1, x2, y2);
        }

        public void EndCappedLine()
        {
            capLinePen = null;
        }

        public void BeginDrawLine(double width, SolidColor color)
        {
            cachePen = new Pen(new SolidColorBrush(color), width);
        }

        public void DrawLine(double x1, double y1, double x2, double y2)
        {
            base.DrawLine(cachePen, new WPFPoint(x1, y1), new WPFPoint(x2, y2));
        }

        public void EndDrawLine()
        {
        }

        public void DrawLeadHeadArrow(Rectangle bounds, SolidColor startColor, SolidColor endColor)
        {
        }

        public Pen GetPen(SolidColor color)
        {
            return resourceManager.GetPen(color);
        }

        public void ReleasePen(Pen pen)
        {
        }

        public Brush GetBrush(SolidColor color)
        {
            return resourceManager.GetBrush(color);
        }

        public void BeginDrawHeaderText(double scale)
        {
            headerTextScale = 9d * scale;
        }

        public void DrawHeaderText(string text, Brush brush, Rectangle rect)
        {
            var ft = new FormattedText(text,
                CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                headerTextTypeface, headerTextScale / 72f * 96f, brush);

            PlatformGraphics.DrawText(ft,
                new Point(rect.X + (rect.Width - ft.Width) / 2, rect.Y + (rect.Height - ft.Height) / 2));
        }

        private static Color DecideTextColor(Cell cell)
        {
            var sheet = cell.Worksheet;
            var controlStyle = sheet.controlAdapter.ControlStyle;
            SolidColor textColor;

            if (!cell.RenderColor.IsTransparent)
                textColor = cell.RenderColor;
            else if (cell.InnerStyle.HasStyle(PlainStyleFlag.TextColor))
                // cell text color, specified by SetRangeStyle
                textColor = cell.InnerStyle.TextColor;
            else if (!controlStyle.TryGetColor(ControlAppearanceColors.GridText, out textColor))
                // default cell text color
                textColor = SolidColor.Black;

            return textColor;
        }

        public RGFont GetFont(string name, double size, FontStyles style)
        {
            return resourceManager.GetTypeface(name, FontWeights.Normal, System.Windows.FontStyles.Normal,
                FontStretches.Normal);
        }
    }

    #endregion // Renderer
}

#endif // WPF