using System.Windows.Media;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Rendering
{
    public interface IRenderer : IGraphics
#if WINFORM
		, System.IDisposable
#endif // WINFORM
    {
        void DrawRunningFocusRect(double x, double y, double w, double h, SolidColor color, int runningOffset);

        void BeginCappedLine(LineCapStyles startCap, Size startSize, LineCapStyles endCap, Size endSize,
            SolidColor color, double width);

        void DrawCappedLine(double x1, double y1, double x2, double y2);

        void EndCappedLine();

        void BeginDrawLine(double width, SolidColor color);

        void DrawLine(double x1, double y1, double x2, double y2);

        void EndDrawLine();

        void DrawCellText(Cell cell, SolidColor textColor, DrawMode drawMode, double scale);

        void UpdateCellRenderFont(Cell cell, UpdateFontReason reason);

        Size MeasureCellText(Cell cell, DrawMode drawMode, double scale);

        void BeginDrawHeaderText(double scale);

        void DrawHeaderText(string text, Brush brush, Rectangle rect);

        void DrawLeadHeadArrow(Rectangle bounds, SolidColor startColor, SolidColor endColor);

        Pen GetPen(SolidColor color);

        void ReleasePen(Pen pen);

        Brush GetBrush(SolidColor color);

        ResourcePoolManager ResourcePoolManager { get; }
    }
}