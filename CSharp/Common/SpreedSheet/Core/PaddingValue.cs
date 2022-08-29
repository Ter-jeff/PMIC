using System;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Padding value struct
    /// </summary>
    [Serializable]
    public struct PaddingValue
    {
        /// <summary>
        ///     Get or set top padding
        /// </summary>
        public double Top { get; set; }

        /// <summary>
        ///     Get or set bottom padding
        /// </summary>
        public double Bottom { get; set; }

        /// <summary>
        ///     Get or set left padding
        /// </summary>
        public double Left { get; set; }

        /// <summary>
        ///     Get or set right padding
        /// </summary>
        public double Right { get; set; }

        /// <summary>
        ///     Create padding and set all values with same specified value.
        /// </summary>
        /// <param name="all">Value applied to all padding.</param>
        public PaddingValue(double all)
            : this(all, all, all, all)
        {
        }

        /// <summary>
        ///     Create padding with every specified values. (in pixel)
        /// </summary>
        /// <param name="top">Top padding.</param>
        /// <param name="bottom">Bottom padding.</param>
        /// <param name="left">Left padding.</param>
        /// <param name="right">Right padding.</param>
        public PaddingValue(double top, double bottom, double left, double right)
            : this()
        {
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
        }

        /// <summary>
        ///     Predefined empty padding value
        /// </summary>
        public static readonly PaddingValue Empty = new PaddingValue(0);

        /// <summary>
        ///     Compare two padding values whether are same
        /// </summary>
        /// <param name="p1">Padding value 1 to be compared</param>
        /// <param name="p2">Padding value 2 to be compared</param>
        /// <returns>True if two padding values are same; otherwise return false</returns>
        public static bool operator ==(PaddingValue p1, PaddingValue p2)
        {
            return p1.Left == p2.Left && p1.Top == p2.Top
                                      && p1.Right == p2.Right && p1.Bottom == p2.Bottom;
        }

        /// <summary>
        ///     Compare two padding values whether are not same
        /// </summary>
        /// <param name="p1">Padding value 1 to be compared</param>
        /// <param name="p2">Padding value 2 to be compared</param>
        /// <returns>True if two padding values are not same; otherwise return false</returns>
        public static bool operator !=(PaddingValue p1, PaddingValue p2)
        {
            return p1.Left != p2.Left || p1.Top != p2.Top
                                      || p1.Right != p2.Right || p1.Bottom != p2.Bottom;
        }

        /// <summary>
        ///     Compare an object and check whether two padding value are same
        /// </summary>
        /// <param name="obj">Another object to be checked</param>
        /// <returns>True if two padding values are same; otherwise return false</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is PaddingValue)) return false;

            var obj2 = (PaddingValue)obj;

            return Top == obj2.Top && Left == obj2.Left
                                   && Right == obj2.Right && Bottom == obj2.Bottom;
        }

        /// <summary>
        ///     Get hash code of this object
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            return (int)(Top + Left * 2 + Right * 3 + Bottom * 4);
        }

        public override string ToString()
        {
            return string.Format("[{0}, {1}, {2}, {3}]", Top, Bottom, Left, Right);
        }
    }
}