#define WPF

namespace SpreedSheet.Interaction
{
    /// <summary>
    ///     Cursor style
    /// </summary>
    public enum CursorStyle : byte
    {
        /// <summary>
        ///     Default (Auto)
        /// </summary>
        PlatformDefault,

        /// <summary>
        ///     Hand
        /// </summary>
        Hand,

        /// <summary>
        ///     Range Selection
        /// </summary>
        Selection,

        /// <summary>
        ///     Full Row Selector
        /// </summary>
        FullRowSelect,

        /// <summary>
        ///     Full Column Selector
        /// </summary>
        FullColumnSelect,

        /// <summary>
        ///     Entire worksheet Selector
        /// </summary>
        EntireSheet,

        /// <summary>
        ///     Move object
        /// </summary>
        Move,

        /// <summary>
        ///     Copy object
        /// </summary>
        Copy,

        /// <summary>
        ///     Change Column Width
        /// </summary>
        ChangeColumnWidth,

        /// <summary>
        ///     Change Row Height
        /// </summary>
        ChangeRowHeight,

        /// <summary>
        ///     Horizontal Resize
        /// </summary>
        ResizeHorizontal,

        /// <summary>
        ///     Vertical Resize
        /// </summary>
        ResizeVertical,

        /// <summary>
        ///     Busy (Waiting)
        /// </summary>
        Busy,

        /// <summary>
        ///     Cross Cursor
        /// </summary>
        Cross
    }
}