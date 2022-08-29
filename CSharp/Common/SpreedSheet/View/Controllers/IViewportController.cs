using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Controllers
{
    internal interface IViewportController : IUserVisual, IVisualController
    {
        Worksheet Worksheet { get; }
        Rectangle Bounds { get; set; }
        IView View { get; }
        IView FocusView { get; set; }
        void Draw(CellDrawingContext dc);
        void UpdateController();
        void Reset();
        void SetViewVisible(ViewTypes view, bool visible);
    }
}