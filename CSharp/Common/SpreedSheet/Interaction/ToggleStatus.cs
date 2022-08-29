using System;

namespace SpreedSheet.Interaction
{
    /// <summary>
    ///     Represent for the button status of mouse
    /// </summary>
    [Flags]
    public enum ToggleStatus
    {
        /// <summary>
        ///     The button has its normal appearance (three-dimensional).
        /// </summary>
        Normal = 0,

        /// <summary>
        ///     The button is inactive (grayed).
        /// </summary>
        Inactive = 0x100,

        /// <summary>
        ///     The button appears pressed.
        /// </summary>
        Pushed = 0x200,

        /// <summary>
        ///     The button has a checked or latched appearance. Use this appearance to show
        ///     that a toggle button has been pressed.
        /// </summary>
        Checked = 0x400,

        /// <summary>
        ///     The button has a flat, two-dimensional appearance.
        /// </summary>
        Flat = 0x4000,

        /// <summary>
        ///     All flags except Normal are set.
        /// </summary>
        All = 0x4700
    }
}