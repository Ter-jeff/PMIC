using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using SpreedSheet.Core.Enum;
using SpreedSheet.Core.Workbook.Appearance;
using unvell.Common;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using FontStyles = unvell.ReoGrid.Drawing.Text.FontStyles;
using Point = System.Windows.Point;
using Size = unvell.ReoGrid.Graphics.Size;

namespace SpreedSheet.Rendering
{
    internal class WpfRenderer : WpfGraphics, IRenderer
    {
        private Pen _cachePen;
        private Pen _capLinePen;
        private double _headerTextScale;
        protected Typeface HeaderTextTypeface;
        private LineCap _lineCap;

        internal WpfRenderer()
        {
            HeaderTextTypeface = new Typeface("Verdana");
            //headerTextTypeface = PlatformUtility.GetFontDefaultTypeface(SystemFonts.SmallCaptionFontFamily);
        }

        public ResourcePoolManager GetResourcePoolManager
        {
            get { return ResourceManager; }
        }

        public Size MeasureCellText(Cell cell, DrawMode drawMode, double scale)
        {
            if (cell.InnerStyle.RotationAngle != 0)
            {
                var m = Matrix.Identity;
                double hw = cell.formattedText.Width * 0.5, hh = cell.formattedText.Height * 0.5;
                Point p1 = new Point(-hw, -hh), p2 = new Point(hw, hh);
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
                    new Point(-cell.formattedText.Width * 0.5, -cell.formattedText.Height * 0.5));
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
                    ResourceManager.GetTypeface(cell.InnerStyle.FontName),
                    fontSize, ResourceManager.GetBrush(textColor));
            }
            else if (reason == UpdateFontReason.FontChanged || reason == UpdateFontReason.ScaleChanged)
            {
                cell.formattedText.SetFontFamily(cell.InnerStyle.FontName);
                cell.formattedText.SetFontSize(fontSize);
            }
            else if (reason == UpdateFontReason.TextColorChanged)
            {
                SolidColor textColor = DecideTextColor(cell);
                cell.formattedText.SetForegroundBrush(ResourceManager.GetBrush(textColor));
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
            _capLinePen = new Pen(new SolidColorBrush(color), width);
            _capLinePen.StartLineCap = PlatformUtility.ToWPFLineCap(startCap);
            _capLinePen.EndLineCap = PlatformUtility.ToWPFLineCap(endCap);
        }

        public void DrawCappedLine(double x1, double y1, double x2, double y2)
        {
            if (_capLinePen != null) base.DrawLine(_capLinePen, x1, y1, x2, y2);
        }

        public void EndCappedLine()
        {
            _capLinePen = null;
        }

        public void BeginDrawLine(double width, SolidColor color)
        {
            _cachePen = new Pen(new SolidColorBrush(color), width);
        }

        public void DrawLine(double x1, double y1, double x2, double y2)
        {
            base.DrawLine(_cachePen, new Point(x1, y1), new Point(x2, y2));
        }

        public void EndDrawLine()
        {
        }

        public void DrawLeadHeadArrow(Rectangle bounds, SolidColor startColor, SolidColor endColor)
        {
        }

        public Pen GetPen(SolidColor color)
        {
            return ResourceManager.GetPen(color);
        }

        public void ReleasePen(Pen pen)
        {
        }

        public Brush GetBrush(SolidColor color)
        {
            return ResourceManager.GetBrush(color);
        }

        public void BeginDrawHeaderText(double scale)
        {
            _headerTextScale = 9d * scale;
        }

        public void DrawHeaderText(string text, Brush brush, Rectangle rect)
        {
            var ft = new FormattedText(text,
                CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                HeaderTextTypeface, _headerTextScale / 72f * 96f, brush);

            PlatformGraphics.DrawText(ft, new unvell.ReoGrid.Graphics.Point(
                rect.X + (rect.Width - ft.Width) / 2, rect.Y + (rect.Height - ft.Height) / 2));
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

        public Typeface GetFont(string name, double size, FontStyles style)
        {
            return ResourceManager.GetTypeface(name, FontWeights.Normal, System.Windows.FontStyles.Normal,
                FontStretches.Normal);
        }
    }
}