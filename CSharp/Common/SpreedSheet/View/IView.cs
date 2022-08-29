using System.Collections.Generic;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View
{
    /// <summary>
    ///     A view is a visual region which can be independent rendered.
    ///     A view can contains multiple child views.
    /// </summary>
    internal interface IView : IUserVisual
    {
        IViewportController ViewportController { get; set; }
        Rectangle Bounds { get; set; }
        double Left { get; }
        double Top { get; }
        double Width { get; }
        double Height { get; }
        double Right { get; }
        double Bottom { get; }
        double ScaleFactor { get; set; }
        bool PerformTransform { get; set; }
        bool Visible { get; set; }
        IList<IView> Children { get; set; }
        void UpdateView();
        void Draw(CellDrawingContext dc);
        void DrawChildren(CellDrawingContext dc);
        Point PointToView(Point p);
        Point PointToController(Point p);
        IView GetViewByPoint(Point p);
    }
}