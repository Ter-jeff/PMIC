using SpreedSheet.Control;

namespace SpreedSheet.Interface
{
    internal interface IMultiSheetAdapter
    {
        ISheetTabControl SheetTabControl { get; }
    }
}