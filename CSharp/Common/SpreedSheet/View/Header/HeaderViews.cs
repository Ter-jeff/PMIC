using SpreedSheet.View.Controllers;

namespace SpreedSheet.View.Header
{
    internal class HeaderView : Viewport, IRangeSelectableView
    {
        protected double HeaderAdjustBackup = 0;

        public HeaderView(IViewportController vc)
            : base(vc)
        {
        }
    }
}