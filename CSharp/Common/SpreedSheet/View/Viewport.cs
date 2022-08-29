using System.Diagnostics;
using SpreedSheet.Enum;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View
{
    internal abstract class Viewport : View, IViewport
    {
        protected Worksheet Sheet;

        public Viewport(IViewportController vc)
            : base(vc)
        {
            Sheet = vc.Worksheet;
        }

        internal Worksheet Worksheet
        {
            get { return Sheet; }
        }

        #region Visible Region

        private GridRegion _visibleRegion;

        /// <summary>
        ///     View window for cell address, decides how many cells are visible for this viewport.
        /// </summary>
        public virtual GridRegion VisibleRegion
        {
            get { return _visibleRegion; }
            set { _visibleRegion = value; }
        }

        #endregion

        #region View window

        private Point _viewStart;

        /// <summary>
        ///     View window start position. (Scroll position)
        /// </summary>
        public virtual Point ViewStart
        {
            get { return _viewStart; }
            set { _viewStart = value; }
        }

        /// <summary>
        ///     Top position of view window. (Vertial scroll position)
        /// </summary>
        public virtual double ViewTop
        {
            get { return _viewStart.Y; }
            set { _viewStart.Y = value; }
        }

        /// <summary>
        ///     Left position of view window. (Horizontal scroll position)
        /// </summary>
        public virtual double ViewLeft
        {
            get { return _viewStart.X; }
            set { _viewStart.X = value; }
        }
        //public virtual RGFloat ViewRight { get { return viewStart.X  + Bounds.Width / this.ScaleFactor; } }
        //public virtual RGFloat ViewBottom { get { return viewStart.Y + Bounds.Height / this.ScaleFactor; } }

        /// <summary>
        ///     The bounds of view window, starts from scroll position, ends at scroll position + window size.
        /// </summary>
        public virtual Rectangle ViewBounds
        {
            get
            {
                return new Rectangle(ScrollViewLeft, ScrollViewTop, Bounds.Width / ScaleFactor,
                    Bounds.Height / ScaleFactor);
            }
        }

        public virtual ScrollDirection ScrollableDirections { get; set; } = ScrollDirection.None;

        public double ScrollX { get; set; }
        public double ScrollY { get; set; }

        public double ScrollViewLeft
        {
            get { return _viewStart.X + ScrollX; }
        }

        public double ScrollViewTop
        {
            get { return _viewStart.Y + ScrollY; }
        }

        public virtual void Scroll(double offX, double offY)
        {
            ScrollX += offX;
            ScrollY += offY;

            if (ScrollX < 0) ScrollX = 0;
            if (ScrollY < 0) ScrollY = 0;
        }

        public virtual void ScrollTo(double x, double y)
        {
            if (x >= 0 && (ScrollableDirections & ScrollDirection.Horizontal) == ScrollDirection.Horizontal)
                ScrollX = x;
            if (y >= 0 && (ScrollableDirections & ScrollDirection.Vertical) == ScrollDirection.Vertical) ScrollY = y;

            if (ScrollX < 0) ScrollX = 0;
            if (ScrollY < 0) ScrollY = 0;
        }

        #endregion // View window

        #region Point transform

        public override Point PointToView(Point p)
        {
            return new Point(
                (p.X - Bounds.Left + ScrollViewLeft * ScaleFactor) / ScaleFactor,
                (p.Y - Bounds.Top + ScrollViewTop * ScaleFactor) / ScaleFactor);
        }

        public override Point PointToController(Point p)
        {
            return new Point(
                (p.X - ScrollViewLeft) * ScaleFactor + Bounds.Left,
                (p.Y - ScrollViewTop) * ScaleFactor + Bounds.Top);
        }

        #endregion // Point transform

        #region Draw

        public override void Draw(CellDrawingContext dc)
        {
#if DEBUG
            var sw = Stopwatch.StartNew();
#endif

            if (!Visible //|| visibleGridRegion == GridRegion.Empty
                || Bounds.Width <= 0 || Bounds.Height <= 0) return;

            //bool needClip = this.Parent == null
            //	|| this.bounds != this.Parent.Bounds;

            //bool needTranslate = this.Parent == null
            //	|| this.viewStart.X != this.Parent.ViewLeft
            //	|| this.ViewStart.Y != this.Parent.ViewTop;

            var g = dc.Graphics;

            if (PerformTransform)
            {
                g.PushClip(Bounds);
                g.PushTransform();
                g.TranslateTransform(Bounds.Left - ScrollViewLeft * ScaleFactor,
                    Bounds.Top - ScrollViewTop * ScaleFactor);
            }

            DrawView(dc);

            if (PerformTransform)
            {
                g.PopTransform();
                g.PopClip();
            }

#if VP_DEBUG
#if WINFORM
			if (this is SheetViewport
				|| this is ColumnHeaderView
				//|| this is RowHeaderView
				|| this is RowOutlineView)
			{
				//var rect = this.bounds;
				//rect.Width--;
				//rect.Height--;
				//dc.Graphics.DrawRectangle(this.bounds, this is SheetViewport ? SolidColor.Blue : SolidColor.Purple);

				var msg = $"{ this.GetType().Name }\n" +
					$"{visibleRegion.ToRange()}\n" +
					$"{this.ViewLeft}, {this.ViewTop}, ({ScrollX}, {ScrollY}), {this.Width}, {this.Height}\n" +
					$"{this.ScrollableDirections}";

				dc.Graphics.PlatformGraphics.DrawString(msg,
						System.Drawing.SystemFonts.DefaultFont, System.Drawing.Brushes.Blue, this.Left + Width / 2, Top + Height / 2);
			}
#elif WPF
			var msg = string.Format("VR {0},{1}-{2},{3} VS X{4},Y{5}\nSD {6}", this.visibleRegion.startRow,
				this.visibleRegion.startCol, this.visibleRegion.endRow, this.visibleRegion.endCol, this.ViewLeft, this.ViewTop,
				this.ScrollableDirections.ToString());

			var ft =
 new System.Windows.Media.FormattedText(msg, System.Globalization.CultureInfo.CurrentCulture, System.Windows.FlowDirection.LeftToRight, 
				new System.Windows.Media.Typeface("Arial"), 12, System.Windows.Media.Brushes.Blue, 96);

			dc.Graphics.PlatformGraphics.DrawText(ft, new System.Windows.Point(this.Left + 1, this.Top + ((this is CellsViewport) ? 30 : this.Height / 2)));
#endif // WPF
#endif // VP_DEBUG

#if DEBUG
            sw.Stop();
            if (sw.ElapsedMilliseconds > 20)
                Debug.WriteLine("draw viewport takes " + sw.ElapsedMilliseconds + " ms. visible region: rows: " +
                                _visibleRegion.Rows + ", cols: " + _visibleRegion.Cols);
#endif // Debug
        }

        public virtual void DrawView(CellDrawingContext dc)
        {
            DrawChildren(dc);
        }

        #endregion // Draw
    }
}