#define WPF


#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF || iOS
using RGFloat = System.Double;
#endif // WPF || iOS
using System;
using System.Globalization;
using System.Windows;

namespace unvell.ReoGrid.Graphics
{
    /// <summary>
    ///     Represents a rectangle that contains x, y, width and height information.
    /// </summary>
    [Serializable]
    public struct Rectangle
    {
        /// <summary>
        ///     Get or set the x-coordinate of rectangle.
        /// </summary>
        public double X;

        /// <summary>
        ///     Get or set the y-coordinate of rectangle.
        /// </summary>
        public double Y;

        /// <summary>
        ///     Get or set the width of rectangle.
        /// </summary>
        public double Width;

        /// <summary>
        ///     Get or set the height of rectangle.
        /// </summary>
        public double Height;

        /// <summary>
        ///     Create rectangle with specified x, y, width and height.
        /// </summary>
        /// <param name="x">Value on x-coordinate of rectangle.</param>
        /// <param name="y">Value on y-coordinate of rectangle.</param>
        /// <param name="width">Width of rectangle.</param>
        /// <param name="height">Height of rectangle.</param>
        public Rectangle(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>
        ///     Create rectangle with specified position and size.
        /// </summary>
        /// <param name="position">Position of rectangle.</param>
        /// <param name="size">Size of rectangle.</param>
        public Rectangle(Point position, Size size)
        {
            X = position.X;
            Y = position.Y;
            Width = size.Width;
            Height = size.Height;
        }

        /// <summary>
        ///     Create rectangle from specified two positions. This method will find the
        ///     top-left position and bottom-right position from two positions and create
        ///     rectangle at correct position.
        /// </summary>
        /// <param name="firstPosition">First position.</param>
        /// <param name="secondPosition">Second position.</param>
        public Rectangle(Point firstPosition, Point secondPosition)
        {
            X = Math.Min(firstPosition.X, secondPosition.X);
            Y = Math.Min(firstPosition.Y, secondPosition.Y);

            var x2 = Math.Max(firstPosition.X, secondPosition.X);
            var y2 = Math.Max(firstPosition.Y, secondPosition.Y);

            Width = x2 - X;
            Height = y2 - Y;
        }

        /// <summary>
        ///     Get or set the location of rectangle.
        /// </summary>
        public Point Location
        {
            get { return new Point(X, Y); }
            set
            {
                X = value.X;
                Y = value.Y;
            }
        }

        /// <summary>
        ///     Get or set the size of rectangle.
        /// </summary>
        public Size Size
        {
            get { return new Size(Width, Height); }
            set
            {
                Width = value.Width;
                Height = value.Height;
            }
        }

        /// <summary>
        ///     Get or set the left position of rectangle.
        /// </summary>
        public double Left
        {
            get { return X; }
        }

        /// <summary>
        ///     Get or set the right position of rectangle.
        /// </summary>
        public double Right
        {
            get { return X + Width; }
        }

        /// <summary>
        ///     Get or set the top position of rectangle.
        /// </summary>
        public double Top
        {
            get { return Y; }
        }

        /// <summary>
        ///     Get or set the bottom position of rectangle.
        /// </summary>
        public double Bottom
        {
            get { return Y + Height; }
        }

        /// <summary>
        ///     Get origin X-coordinate of rectangle.
        /// </summary>
        public double OriginX
        {
            get { return X + Width / 2; }
        }

        /// <summary>
        ///     Get origin Y-coordinate of rectangle.
        /// </summary>
        public double OriginY
        {
            get { return Y + Height / 2; }
        }

        /// <summary>
        ///     Get origin of rectangle.
        /// </summary>
        public Point Origin
        {
            get { return new Point(X + Width / 2, Y + Height / 2); }
        }

        /// <summary>
        ///     Check whether or not the specified point is contained by this rectangle.
        /// </summary>
        /// <param name="p">Point to be checked.</param>
        /// <returns>True if the point is contained by this rectangle; Otherwise return false;</returns>
        public bool Contains(Point p)
        {
            return p.X >= X && p.Y >= Y && p.X <= Right && p.Y <= Bottom;
        }

        /// <summary>
        ///     Check whether or not the specified point (described by x and y) is contained by this rectangle.
        /// </summary>
        /// <param name="x">Value on x-coordinate.</param>
        /// <param name="y">Value on y-coordinate.</param>
        /// <returns>True if the point is contained by this rectangle; Otherwise return false;</returns>
        public bool Contains(double x, double y)
        {
            return x >= X && y >= Y && x <= Right && y <= Bottom;
        }

        /// <summary>
        ///     Move the rectangle by amount specified by x and y coordinates.
        /// </summary>
        /// <param name="x">Value on x-coordinate.</param>
        /// <param name="y">Value on y-coordinate.</param>
        public void Offset(double x, double y)
        {
            X += x;
            Y += y;
        }

        /// <summary>
        ///     Inflate the rectangle by amount specified by x and y coordinates.
        ///     <remarks>
        ///         It is also possible to shrink this rectangle by specifying negative values.
        ///     </remarks>
        /// </summary>
        /// <param name="x">Value on x-coordinate.</param>
        /// <param name="y">Value on y-coordinate.</param>
        public void Inflate(double x, double y)
        {
            X -= x;
            Y -= y;
            Width += x + x;
            Height += y + y;
        }

        /// <summary>
        ///     Determines if this rectangle intersets with rect.
        /// </summary>
        /// <param name="rect">The rectangle to test.</param>
        /// <returns>This method returns true if there is any intersection, otherwise false.</returns>
        public bool IntersectWith(Rectangle rect)
        {
            return rect.X < X + Width &&
                   X < rect.X + rect.Width &&
                   rect.Y < Y + Height &&
                   Y < rect.Y + rect.Height;
        }

        /// <summary>
        ///     Creates a Rectangle that represents the intersection between this Rectangle and rect.
        /// </summary>
        /// <param name="rect">The rectangle to test.</param>
        public void Intersect(Rectangle rect)
        {
            var result = Intersect(rect, this);

            X = result.X;
            Y = result.Y;
            Width = result.Width;
            Height = result.Height;
        }

        /// <summary>
        ///     Check two rectangles and calculate the intersection of two rectangles.
        ///     If no intersection detected, a rectangle with zero width and height is returned.
        /// </summary>
        /// <param name="a">First rectangle to be test.</param>
        /// <param name="b">Second rectangle to be test.</param>
        /// <returns>Intersected rectangle.</returns>
        public static Rectangle Intersect(Rectangle a, Rectangle b)
        {
            var x1 = Math.Max(a.X, b.X);
            var x2 = Math.Min(a.X + a.Width, b.X + b.Width);
            var y1 = Math.Max(a.Y, b.Y);
            var y2 = Math.Min(a.Y + a.Height, b.Y + b.Height);

            if (x2 >= x1 && y2 >= y1)
                return new Rectangle(x1, y1, x2 - x1, y2 - y1);
            return new Rectangle();
        }

        /// <summary>
        ///     Compare two rectangles to check whether or not they are same.
        /// </summary>
        /// <param name="obj">Another rectange compared to this rectangle.</param>
        /// <returns>True if two rectangles are same; Otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Rectangle)) return false;

            var rect2 = (Rectangle)obj;

            return X == rect2.X && Y == rect2.Y && Width == rect2.Width && Height == rect2.Height;
        }

        /// <summary>
        ///     Get hash code of this rectangle.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return (int)((uint)X ^
                         (((uint)Y << 13) | ((uint)Y >> 19)) ^
                         (((uint)Width << 26) | ((uint)Width >> 6)) ^
                         (((uint)Height << 7) | ((uint)Height >> 25)));
        }

        //public override string ToString()
        //{
        //	return string.Format("Rectangle[{0}, {1}, {2}, {3}]", this.X, this.Y, this.Width, this.Height);
        //}

        /// <summary>
        ///     Scale rectangle by multiplying specified scale factor.
        /// </summary>
        /// <param name="r">The rectangle to be scaled.</param>
        /// <param name="s">Scale factor to be multiplied.</param>
        /// <returns></returns>
        public static Rectangle operator *(Rectangle r, double s)
        {
            return new Rectangle(r.X * s, r.Y * s, r.Width * s, r.Height * s);
        }

        /// <summary>
        ///     Compare two rectangels to check whether or not they are same.
        /// </summary>
        /// <param name="r1">First rectangle to be compared.</param>
        /// <param name="r2">Second rectangle to be compared.</param>
        /// <returns>True if two rectangles are same; Otherwise return false.</returns>
        public static bool operator ==(Rectangle r1, Rectangle r2)
        {
            return r1.X == r2.X && r1.Y == r2.Y && r1.Width == r2.Width && r1.Height == r2.Height;
        }

        /// <summary>
        ///     Compare two rectangles to check whether or not they are not same.
        /// </summary>
        /// <param name="r1">First rectangle to be compared.</param>
        /// <param name="r2">Second rectangle to be compared.</param>
        /// <returns>True if two rectangles are not same; Otherwise return false.</returns>
        public static bool operator !=(Rectangle r1, Rectangle r2)
        {
            return r1.X != r2.X || r1.Y != r2.Y || r1.Width != r2.Width || r1.Height != r2.Height;
        }

        /// <summary>
        ///     Convert this rectangle into string. (Format: {x, y, w, h})
        /// </summary>
        /// <returns>String converted from this rectangle object.</returns>
        public override string ToString()
        {
            return string.Format("{{{0}, {1}, {2}, {3}}}", X.ToString(CultureInfo.CurrentCulture),
                Y.ToString(CultureInfo.CurrentCulture),
                Width.ToString(CultureInfo.CurrentCulture),
                Height.ToString(CultureInfo.CurrentCulture));
        }

        #region Platform Associated

#if WINFORM
		/// <summary>
		/// Convert System.Drawing.Rectangle to unvell.ReoGrid.Graphics.Rectangle.
		/// </summary>
		/// <param name="r">Rectangle of System.Drawing.Rectangle.</param>
		/// <returns>Rectangle of unvell.ReoGrid.Graphics.Rectangle.</returns>
		public static implicit operator Rectangle(System.Drawing.Rectangle r)
		{
			return new Rectangle(r.X, r.Y, r.Width, r.Height);
		}
		/// <summary>
		/// Convert System.Drawing.RectangleF to unvell.ReoGrid.Graphics.Rectangle.
		/// </summary>
		/// <param name="r">Rectangle of System.Drawing.RectangleF.</param>
		/// <returns>Rectangle of unvell.ReoGrid.Graphics.Rectangle.</returns>
		public static implicit operator Rectangle(System.Drawing.RectangleF r)
		{
			return new Rectangle(r.X, r.Y, r.Width, r.Height);
		}
		/// <summary>
		/// Convert unvell.ReoGrid.Graphics.Rectangle to System.Drawing.Rectangle.
		/// </summary>
		/// <param name="r">Rectangle of unvell.ReoGrid.Graphics.Rectangle.</param>
		/// <returns>Rectangle of System.Drawing.Rectangle.</returns>
		public static explicit operator System.Drawing.Rectangle(Rectangle r)
		{
			return new System.Drawing.Rectangle((int)Math.Round(r.X), (int)Math.Round(r.Y), (int)Math.Round(r.Width), (int)Math.Round(r.Height));
		}
		/// <summary>
		/// Convert unvell.ReoGrid.Graphics.Rectangle to System.Drawing.RectangleF.
		/// </summary>
		/// <param name="r">Rectangle of unvell.ReoGrid.Graphics.Rectangle.</param>
		/// <returns>Rectangle of System.Drawing.RectangleF.</returns>
		public static implicit operator System.Drawing.RectangleF(Rectangle r)
		{
			return new System.Drawing.RectangleF(r.X, r.Y, r.Width, r.Height);
		}
#endif // WINFORM

#if ANDROID
		public static implicit operator Android.Graphics.RectF(Rectangle r)
		{
			return new Android.Graphics.RectF(r.Left, r.Top, r.Right, r.Bottom);
		}
		public static implicit operator Rectangle(Android.Graphics.RectF rect)
		{
			return new Rectangle(rect.Left, rect.Top, rect.Width(), rect.Height());
		}
		public static implicit operator Android.Graphics.Rect(Rectangle r)
		{
			return new Android.Graphics.Rect((int)r.Left, (int)r.Top, (int)r.Right, (int)r.Bottom);
		}
		public static implicit operator Rectangle(Android.Graphics.Rect rect)
		{
			return new Rectangle(rect.Left, rect.Top, rect.Width(), rect.Height());
		}
#endif // ANDROID

#if WPF
        public static implicit operator Rect(Rectangle r)
        {
            return new Rect(r.X, r.Y, r.Width, r.Height);
        }

        public static implicit operator Rectangle(Rect r)
        {
            return new Rectangle(r.X, r.Y, r.Width, r.Height);
        }
#endif // WPF

#if iOS
		public static implicit operator CoreGraphics.CGRect(Rectangle r)
		{
			return new CoreGraphics.CGRect(r.X, r.Y, r.Width, r.Height);
		}
		public static implicit operator Rectangle(CoreGraphics.CGRect r)
		{
			return new Rectangle(r.X, r.Y, r.Width, r.Height);
		}
#endif // iOS

        #endregion // Platform Associated
    }
}

/* Performance test
 * 
			StringBuilder sb = new StringBuilder();
			Stopwatch sw = Stopwatch.StartNew();

			Graphics.Rectangle r1= new Graphics.Rectangle();
			Graphics.Rectangle r2=new Graphics.Rectangle();

			for (int i = 0; i < 10000000; i++)
			{
				r1 = new Graphics.Rectangle();
				//r2 = r1;
			}

			sw.Stop();
			sb.AppendLine(sw.ElapsedMilliseconds + " ms.");
			sw.Reset();
			sw.Start();

			System.Drawing.RectangleF rr1 = new System.Drawing.RectangleF();
			System.Drawing.RectangleF rr2 = new System.Drawing.RectangleF();

			for (int i = 0; i < 10000000; i++)
			{
				rr1 = new System.Drawing.RectangleF();
				//rr2 = rr1;
			}

			sw.Stop();
			sb.AppendLine(sw.ElapsedMilliseconds + " ms.");
			sw.Reset();
			sw.Start();

			Graphics.Rectangle rrrr1 = new Graphics.Rectangle();
			Graphics.Rectangle rrrr2 = new Graphics.Rectangle();

			for (int i = 0; i < 10000000; i++)
			{
				rrrr1 = new Graphics.Rectangle(1f, 1f, 1f, 1f);
				//rrrr2 = rrrr1;
			}

			sw.Stop();
			sb.AppendLine(sw.ElapsedMilliseconds + " ms.");
			sw.Reset();
			sw.Start();

			System.Drawing.RectangleF rrrrr1 = new System.Drawing.RectangleF();
			System.Drawing.RectangleF rrrrr2 = new System.Drawing.RectangleF();

			for (int i = 0; i < 10000000; i++)
			{
				rrrrr1 = new System.Drawing.RectangleF(1f, 1f, 1f, 1f);
				//rr2 = rr1;
			}

			sw.Stop();
			sb.AppendLine(sw.ElapsedMilliseconds + " ms.");
			sw.Reset();
			sw.Start();

			System.Drawing.RectangleF rrr1 = new System.Drawing.RectangleF();
			Graphics.Rectangle rrr2 = new Graphics.Rectangle();

			for (int i = 0; i < 10000000; i++)
			{
				rrr1 = new System.Drawing.RectangleF(1, 1, 1, 1);
				rrr2 = rrr1;
			}

			sw.Stop();
			sb.AppendLine(sw.ElapsedMilliseconds + " ms.");

			MessageBox.Show(sb.ToString());
*/