using SpreedSheet.Core;
using unvell.ReoGrid;

namespace SpreedSheet.View.Controllers
{
    /// <summary>
    ///     Interface for freezable ViewportController
    /// </summary>
    internal interface IFreezableViewportController
    {
        /// <summary>
        ///     Freeze to specified cell and position.
        /// </summary>
        /// <param name="pos">Position of cell to start freeze.</param>
        /// <param name="area">Decides the frozen view area.</param>
        void Freeze(CellPosition pos, FreezeArea area = FreezeArea.LeftTop);
    }
}