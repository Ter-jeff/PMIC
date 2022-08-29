using SpreedSheet.View;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Interface
{
    internal interface IShowContextMenuAdapter
    {
        void ShowContextMenuStrip(ViewTypes viewType, Point containerLocation);
    }
}