using SpreedSheet.View.Controllers;

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        internal void InitViewportController()
        {
            ViewportController = new NormalViewportController(this);
        }
    }
}