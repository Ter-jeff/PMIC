using System;

namespace SpreedSheet.Core.Enum
{
    /// <summary>
    ///     Cell horizontal alignment (default: General)
    /// </summary>
    [Serializable]
    public enum GridHorAlign
    {
        /// <summary>
        ///     General horizontal alignment (Spreadsheet decides the alignment automatically)
        /// </summary>
        General,

        /// <summary>
        ///     Left
        /// </summary>
        Left,

        /// <summary>
        ///     Center
        /// </summary>
        Center,

        /// <summary>
        ///     Right
        /// </summary>
        Right,

        /// <summary>
        ///     Distributed to fill the space of cell
        /// </summary>
        DistributedIndent
    }
}