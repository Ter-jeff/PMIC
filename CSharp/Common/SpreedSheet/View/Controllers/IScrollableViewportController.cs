using SpreedSheet.Core;
using SpreedSheet.Enum;

namespace SpreedSheet.View.Controllers
{
    internal interface IScrollableViewportController
    {
        void HorizontalScroll(double value);

        void VerticalScroll(double value);

        void ScrollViews(ScrollDirection dir, double x, double y);

        void ScrollOffsetViews(ScrollDirection dir, double x, double y);

        void ScrollToRange(RangePosition range, CellPosition pos);

        void SynchronizeScrollBar();
    }
}