using SpreedSheet.Enum;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View
{
    internal interface IViewport : IView
    {
        Point ViewStart { get; set; }
        double ViewTop { get; }

        double ViewLeft { get; }
        //RGFloat ViewRight { get; }
        //RGFloat ViewBottom { get; }

        double ScrollX { get; set; }
        double ScrollY { get; set; }
        double ScrollViewTop { get; }
        double ScrollViewLeft { get; }

        ScrollDirection ScrollableDirections { get; set; }

        GridRegion VisibleRegion { get; set; }
        void Scroll(double offX, double offY);

        void ScrollTo(double x, double y);
    }
}