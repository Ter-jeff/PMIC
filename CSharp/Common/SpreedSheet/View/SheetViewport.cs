using System.Collections.Generic;
using SpreedSheet.View.Controllers;

namespace SpreedSheet.View
{
    internal class SheetViewport : LayerViewport
    {
        public SheetViewport(IViewportController vc)
            : base(vc)
        {
            Children = new List<IView>(4)
            {
                new CellsViewport(vc) { PerformTransform = false },
                new CellsForegroundView(vc) { PerformTransform = false }
            };
        }
    }
}