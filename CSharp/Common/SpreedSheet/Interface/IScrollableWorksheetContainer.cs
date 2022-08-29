using unvell.ReoGrid;

namespace SpreedSheet.Interface
{
    public interface IScrollableWorksheetContainer
    {
        bool ShowScrollEndSpacing { get; }
        void RaiseWorksheetScrolledEvent(Worksheet worksheet, double x, double y);
    }
}