using SpreedSheet.View.Controllers;

namespace SpreedSheet.View
{
    internal abstract class LayerViewport : Viewport
    {
        public LayerViewport(IViewportController vc)
            : base(vc)
        {
        }

        public override void UpdateView()
        {
            if (Children != null)
                foreach (var child in Children)
                {
                    child.Bounds = Bounds;
                    child.ScaleFactor = ScaleFactor;

                    var childViewport = child as IViewport;
                    if (childViewport != null)
                    {
                        childViewport.ViewStart = ViewStart;
                        childViewport.ScrollX = ScrollX;
                        childViewport.ScrollY = ScrollY;
                        childViewport.VisibleRegion = VisibleRegion;
                        childViewport.ScrollableDirections = ScrollableDirections;
                    }

                    child.UpdateView();
                }
        }
    }
}