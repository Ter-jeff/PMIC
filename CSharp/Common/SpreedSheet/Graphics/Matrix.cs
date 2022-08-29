#define WPF


#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using System;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WPF

namespace unvell.ReoGrid.Graphics
{
    #region Matrix

    /// <summary>
    ///     Matrix for 2D graphics.
    /// </summary>
    [Serializable]
    public class Matrix3x2f
    {
        /// <summary>
        ///     Predefined identify matrix.
        /// </summary>
        public static readonly Matrix3x2f Identify = new Matrix3x2f
        {
            a1 = 1, b1 = 0,
            a2 = 0, b2 = 1,
            a3 = 0, b3 = 0
        };

        /// <summary>
        ///     Translate this matrix.
        /// </summary>
        /// <param name="x">value of x-coordinate to be offset.</param>
        /// <param name="y">Value of y-coordinate to be offset.</param>
        public void Translate(double x, double y)
        {
            a3 += x;
            b3 += y;
        }

        /// <summary>
        ///     Rotate this matrix.
        /// </summary>
        /// <param name="angle">Angle to be rotated.</param>
        public void Rotate(float angle)
        {
            var radians = angle / 180f * Math.PI;
            var sin = Math.Sin(radians);
            var cos = Math.Cos(radians);

            a1 = cos;
            b1 = sin;
            a2 = -sin;
            b2 = cos;
        }

        /// <summary>
        ///     Scale this matrix.
        /// </summary>
        /// <param name="x">Value of x-aspect to be scaled.</param>
        /// <param name="y">Value of y-aspect to be scaled.</param>
        public void Scale(float x, float y)
        {
            a1 *= x;
            b1 *= x;
            a2 *= y;
            b2 *= y;
        }
#pragma warning disable 1591
        public double a1, b1;
        public double a2, b2;
        public double a3, b3;
#pragma warning restore 1591
    }

    #endregion // Matrix
}