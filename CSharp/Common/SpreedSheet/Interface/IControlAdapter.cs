namespace SpreedSheet.Interface
{
    internal interface IControlAdapter : ICompViewAdapter,
        IEditableControlAdapter, IScrollableControlAdapter, ITimerSupportedAdapter,
        IShowContextMenuAdapter
    {
    }
}