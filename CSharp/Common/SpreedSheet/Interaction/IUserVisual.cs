using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Interaction
{
    /// <summary>
    ///     Represents an user-interactive object in ReoGrid cross-platform views system.
    /// </summary>
    public interface IUserVisual
    {
        /// <summary>
        ///     Handle mouse down event
        /// </summary>
        /// <param name="location">Transformed relative location to this object</param>
        /// <param name="buttons">Current mouse button pressing status</param>
        /// <returns>True if event handled; otherwise return false</returns>
        bool OnMouseDown(Point location, MouseButtons buttons);

        /// <summary>
        ///     Handle mouse move event
        /// </summary>
        /// <param name="location">Transformed relative location to this object</param>
        /// <param name="buttons">Current mouse button pressing status</param>
        /// <returns>True if event handled; otherwise return false</returns>
        bool OnMouseMove(Point location, MouseButtons buttons);

        /// <summary>
        ///     Handle mouse up event
        /// </summary>
        /// <param name="location">Transformed relative location to this object</param>
        /// <param name="buttons">Current mouse button pressing status</param>
        /// <returns>True if event handled; otherwise return false</returns>
        bool OnMouseUp(Point location, MouseButtons buttons);

        /// <summary>
        ///     Handle mouse double click event
        /// </summary>
        /// <param name="location">Transformed relative location to this object</param>
        /// <param name="buttons">Current mouse button pressing status</param>
        /// <returns>True if event handled; otherwise return false</returns>
        bool OnMouseDoubleClick(Point location, MouseButtons buttons);

        /// <summary>
        ///     Handle key down event
        /// </summary>
        /// <param name="keys">ReoGrid virtual keys (equal to System.Windows.Forms.Keys)</param>
        /// <returns>True if event handled; otherwise return false</returns>
        bool OnKeyDown(KeyCode keys);

        /// <summary>
        ///     Set this object to get user interface focus. Object after get focus can always
        ///     receive user's mouse and keyboard input.
        /// </summary>
        void SetFocus();

        /// <summary>
        ///     Release user interface focus from this object. This object will no longer be able to
        ///     receive user's mouse and keyboard input.
        /// </summary>
        void FreeFocus();

        /// <summary>
        ///     Redraw this object.
        /// </summary>
        void Invalidate();
    }
}