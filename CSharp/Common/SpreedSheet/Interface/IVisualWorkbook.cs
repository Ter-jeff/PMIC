using unvell.ReoGrid;

namespace SpreedSheet.Interface
{
    public interface IVisualWorkbook : IScrollableWorksheetContainer
    {
        Worksheet ActiveWorksheet { get; set; }
    }
}