using System;

namespace SpreedSheet.Core.Enum
{
    /// <summary>
    ///     Cell vertical alignment (default: Middle)
    /// </summary>
    [Serializable]
    public enum GridVerAlign
    {
        /// <summary>
        ///     Default
        /// </summary>
        General,

        /// <summary>
        ///     Top
        /// </summary>
        Top,

        /// <summary>
        ///     Middle
        /// </summary>
        Middle,

        /// <summary>
        ///     Bottom
        /// </summary>
        Bottom
    }
}