using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Interface
{
    internal interface ICompViewAdapter : IMultiSheetAdapter
    {
        IVisualWorkbook ControlInstance { get; }
        IRenderer Renderer { get; }
        ControlAppearanceStyle ControlStyle { get; }
        double BaseScale { get; }
        double MinScale { get; }
        double MaxScale { get; }
        bool IsVisible { get; }
        void ChangeCursor(CursorStyle cursor);
        void RestoreCursor();
        void ChangeSelectionCursor(CursorStyle cursor);
        Rectangle GetContainerBounds();
        void Focus();
        void Invalidate();
        void ChangeBackgroundColor(SolidColor color);
        Point PointToScreen(Point point);
        void ShowTooltip(Point point, string content);
    }
}