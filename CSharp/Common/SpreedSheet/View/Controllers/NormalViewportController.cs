using System;
using System.Diagnostics;
using System.Linq;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using SpreedSheet.Enum;
using SpreedSheet.Rendering;
using SpreedSheet.View.Header;
using unvell.ReoGrid;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Controllers
{
    /// <summary>
    ///     Standard view controller for normal scene of control
    /// </summary>
    internal class NormalViewportController : ViewportController, IFreezableViewportController,
        IScrollableViewportController, IScalableViewportController
    {
        #region Constructor

        private readonly LeadHeaderView _leadHeadPart;

        // four-header (two columns, two rows)
        private RowHeaderView _rowHeaderPart1;
        private readonly RowHeaderView _rowHeaderPart2;
        private readonly ColumnHeaderView _colHeaderPart2;
        private ColumnHeaderView _colHeaderPart1;

        // four-viewport, any of one could be set as main-viewport, and others will be frozen
        private SheetViewport _topLeftViewport;
        private SheetViewport _leftBottomViewport;
        private readonly SheetViewport _rightBottomViewport;
        private SheetViewport _rightTopViewport;

        // main-viewport used to decide the value of scrollbar
        internal SheetViewport MainViewport;

        public NormalViewportController(Worksheet sheet)
            : base(sheet)
        {
            sheet.ViewportController = this;

            // unfrozen
            AddView(_leadHeadPart = new LeadHeaderView(this));
            AddView(_colHeaderPart2 = new ColumnHeaderView(this));
            AddView(_rowHeaderPart2 = new RowHeaderView(this));
            AddView(_rightBottomViewport = new SheetViewport(this));

            // default settings of viewports
            _rightBottomViewport.ScrollableDirections = ScrollDirection.Both;

            FocusView = _rightBottomViewport;
            MainViewport = _rightBottomViewport;
        }

        #endregion // Constructor

        #region Visibility Management

        public override void SetViewVisible(ViewTypes viewFlag, bool visible)
        {
            base.SetViewVisible(viewFlag, visible);

            var rowHeadVisible = IsViewVisible(ViewTypes.RowHeader);
            var colHeadVisible = IsViewVisible(ViewTypes.ColumnHeader);

            var rowOutlineHeadVisible = IsViewVisible(ViewTypes.RowOutline | ViewTypes.ColumnHeader);
            var outlineColHeadVisible = IsViewVisible(ViewTypes.ColOutline | ViewTypes.RowHeader);

            #region Column Head

            if ((viewFlag & ViewTypes.ColumnHeader) == ViewTypes.ColumnHeader)
            {
                if (visible)
                {
                    _colHeaderPart2.Visible = true;
                    _colHeaderPart2.VisibleRegion = _rightBottomViewport.VisibleRegion;

                    if (_currentFreezePos.Col > 0)
                    {
                        _colHeaderPart1.Visible = true;
                        _colHeaderPart1.VisibleRegion = _frozenVisibleRegion;
                    }
                }
                else
                {
                    _colHeaderPart2.Visible = false;

                    if (_colHeaderPart1 != null) _colHeaderPart1.Visible = false;
                }
            }

            #endregion // Column Head

            #region Row Head

            if ((viewFlag & ViewTypes.RowHeader) == ViewTypes.RowHeader)
            {
                if (visible)
                {
                    _rowHeaderPart2.Visible = true;
                    _rowHeaderPart2.VisibleRegion = _rightBottomViewport.VisibleRegion;

                    if (_currentFreezePos.Row > 0)
                    {
                        _rowHeaderPart1.Visible = true;
                        _rowHeaderPart1.VisibleRegion = _frozenVisibleRegion;
                    }
                }
                else
                {
                    _rowHeaderPart2.Visible = false;

                    if (_rowHeaderPart1 != null) _rowHeaderPart1.Visible = false;
                }
            }

            #endregion

            _leadHeadPart.Visible = IsViewVisible(ViewTypes.LeadHeader);

#if OUTLINE
			bool rowOutlineVisible = IsViewVisible(ViewTypes.RowOutline);
			bool colOutlineVisible = IsViewVisible(ViewTypes.ColOutline);

			// row outline
			if (rowOutlineVisible)
			{
				if (visible && this.rowOutlinePart2 == null)
				{
					this.rowOutlinePart2 = new RowOutlineView(this);

					// set view start position
					this.rowOutlinePart2.ScrollY = this.rowHeaderPart2.ScrollY;

					this.AddView(this.rowOutlinePart2);
				}

				if (rowOutlinePart2 != null) rowOutlinePart2.Visible = rowOutlineVisible;
				if (rowOutlinePart1 != null) rowOutlinePart1.Visible = rowOutlineVisible;
			}
			else
			{
				if (rowOutlinePart2 != null) rowOutlinePart2.Visible = false;
				if (rowOutlinePart1 != null) rowOutlinePart1.Visible = false;
			}

			// row outline header
			if (rowOutlineHeadVisible)
			{
				if (this.rowOutlineHeadPart == null)
				{
					this.rowOutlineHeadPart = new RowOutlineHeaderView(this);
					this.AddView(this.rowOutlineHeadPart);
				}

				this.rowOutlineHeadPart.Visible = true;
			}
			else if (this.rowOutlineHeadPart != null)
			{
				this.rowOutlineHeadPart.Visible = false;
			}

			// column outline
			if (colOutlineVisible)
			{
				if (visible && this.colOutlinePart2 == null)
				{
					this.colOutlinePart2 = new ColumnOutlinePart(this);

					// set view start position
					this.colOutlinePart2.ScrollX = this.colHeaderPart2.ScrollX;

					this.AddView(this.colOutlinePart2);
				}

				if (colOutlinePart2 != null) colOutlinePart2.Visible = colOutlineVisible;
				if (colOutlinePart1 != null) colOutlinePart1.Visible = colOutlineVisible;
			}
			else
			{
				if (colOutlinePart2 != null) colOutlinePart2.Visible = false;
				if (colOutlinePart1 != null) colOutlinePart1.Visible = false;
			}

			// column outline header
			if (outlineColHeadVisible)
			{
				if (this.colOutlineHeadPart == null)
				{
					this.colOutlineHeadPart = new ColumnOutlineHeadPart(this);
					this.AddView(this.colOutlineHeadPart);
				}

				this.colOutlineHeadPart.Visible = true;
			}
			else if (this.colOutlineHeadPart != null)
			{
				this.colOutlineHeadPart.Visible = false;
			}

			// outline space
			if (rowOutlineVisible && colOutlineVisible)
			{
				if (this.outlineLeftTopSpace == null)
				{
					this.outlineLeftTopSpace = new OutlineLeftTopSpace(this);
					this.AddView(outlineLeftTopSpace);
				}

				this.outlineLeftTopSpace.Visible = visible;
			}
			else if (this.outlineLeftTopSpace != null)
			{
				this.outlineLeftTopSpace.Visible = false;
			}
#endif // OUTLINE
        }

        private Rectangle GetGridScaleBounds(CellPosition pos)
        {
            return GetGridScaleBounds(pos.Row, pos.Col);
        }

        private Rectangle GetGridScaleBounds(int row, int col)
        {
            var freezePos = Worksheet.FreezePos;

            var rowHead = Worksheet.rows[freezePos.Row];
            var colHead = Worksheet.cols[freezePos.Col];

            return new Rectangle(colHead.Left * ScaleFactor, rowHead.Top * ScaleFactor,
                colHead.InnerWidth * ScaleFactor + 1, rowHead.InnerHeight * ScaleFactor + 1);
        }

        private Rectangle GetRangeScaledBounds(GridRegion region)
        {
            var startRowHead = Worksheet.rows[region.StartRow];
            var startColHead = Worksheet.cols[region.StartCol];

            var endRowHead = Worksheet.rows[region.EndRow];
            var endColHead = Worksheet.cols[region.EndCol];

            var x1 = startColHead.Left * ScaleFactor;
            var y1 = startRowHead.Top * ScaleFactor;
            var x2 = endColHead.Right * ScaleFactor;
            var y2 = endRowHead.Bottom * ScaleFactor;

            return new Rectangle(x1, y1, x2 - x1, y2 - y1);
        }

        public override IView FocusView
        {
            get { return base.FocusView; }

            set
            {
                base.FocusView = value;

                if (value == null) base.FocusView = _rightBottomViewport;
            }
        }

        #endregion // Visibility Management

        #region Update

        private GridRegion _frozenVisibleRegion = GridRegion.Empty;
        private GridRegion _mainVisibleRegion = GridRegion.Empty;

        protected virtual void UpdateVisibleRegion()
        {
            // update right bottom visible region
            UpdateViewportVisibleRegion(_rightBottomViewport, GetVisibleRegion(_rightBottomViewport));

            var freezePos = Worksheet.FreezePos;

            if (freezePos.Row > 0 || freezePos.Col > 0)
            {
                // update left top visible region
                UpdateViewportVisibleRegion(_topLeftViewport, GetVisibleRegion(_topLeftViewport));

                // update left bottom visible region
                UpdateViewportVisibleRegion(_leftBottomViewport, new GridRegion(
                    _rightBottomViewport.VisibleRegion.StartRow,
                    _topLeftViewport.VisibleRegion.StartCol,
                    _rightBottomViewport.VisibleRegion.EndRow,
                    _topLeftViewport.VisibleRegion.EndCol));

                // update right top visible region
                UpdateViewportVisibleRegion(_rightTopViewport, new GridRegion(
                    _topLeftViewport.VisibleRegion.StartRow,
                    _rightBottomViewport.VisibleRegion.StartCol,
                    _topLeftViewport.VisibleRegion.EndRow,
                    _rightBottomViewport.VisibleRegion.EndCol));
            }

            // frozen headers always are synchronized to left top viewport
            if (_topLeftViewport != null)
            {
                _colHeaderPart1.VisibleRegion = _topLeftViewport.VisibleRegion;
                _rowHeaderPart1.VisibleRegion = _topLeftViewport.VisibleRegion;
            }

            // normal headers always are synchronized to right bottom viewport
            _colHeaderPart2.VisibleRegion = _rightBottomViewport.VisibleRegion;
            _rowHeaderPart2.VisibleRegion = _rightBottomViewport.VisibleRegion;
        }

        private void UpdateViewportVisibleRegion(Viewport viewport, GridRegion range)
        {
            var oldVisible = viewport.VisibleRegion;
            viewport.VisibleRegion = range;

            UpdateNewVisibleRegionTexts(viewport.VisibleRegion, oldVisible);
        }

        // TODO: Need performance improvement
        private void UpdateNewVisibleRegionTexts(GridRegion region, GridRegion oldVisibleRegion)
        {
#if DEBUG
            var sw = Stopwatch.StartNew();
#endif

            // TODO: Need performance improvement
            //       do not perform this during visible region updating
            //
            // end of visible region updating
            Worksheet.cells.Iterate(region.StartRow, region.StartCol, region.Rows, region.Cols, true, (r, c, cell) =>
            {
                var rowHeader = Worksheet.rows[r];
                if (rowHeader.InnerHeight <= 0) return region.Cols;

                int cspan = cell.Colspan;
                if (cspan <= 0) return 1;

                if (cell.RenderScaleFactor != ScaleFactor
                    && !string.IsNullOrEmpty(cell.DisplayText))
                {
                    Worksheet.UpdateCellFont(cell, UpdateFontReason.ScaleChanged);
                    cell.RenderScaleFactor = ScaleFactor;
                }

                return cspan <= 0 ? 1 : cspan;
            });

#if DEBUG
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;

            if (ms > 10) Debug.WriteLine("update new visible region text takes " + ms + " ms.");
#endif // DEBUG
        }

        #region Visible Region Update

        /// <summary>
        ///     Update visible region for viewport. Visible region decides how many rows and columns
        ///     of cells (from...to) will be displayed.
        /// </summary>
        internal static GridRegion GetVisibleRegion(Viewport viewport)
        {
#if DEBUG
            var watch = Stopwatch.StartNew();
#endif

            var sheet = viewport.Worksheet;

            var scale = sheet.renderScaleFactor;

            //Point calcedViewStart = viewport.ViewStart;

            var region = GridRegion.Empty;

            var scaledViewLeft = viewport.ScrollViewLeft;
            var scaledViewTop = viewport.ScrollViewTop;
            var scaledViewRight = viewport.ScrollViewLeft + viewport.Width / scale;
            var scaledViewBottom = viewport.ScrollViewTop + viewport.Height / scale;

            // begin visible region updating
            if (viewport.Height > 0 && sheet.rows.Count > 0)
            {
                float contentBottom = sheet.rows.Last().Bottom;

                if (scaledViewTop > contentBottom)
                {
                    region.StartRow = sheet.RowCount - 1;
                }
                else
                {
                    var index = sheet.rows.Count / 2;

                    ArrayHelper.QuickFind(index, 0, sheet.rows.Count, rindex =>
                    {
                        var row = sheet.rows[rindex];

                        float top = row.Top;
                        float bottom = row.Bottom;

                        if (scaledViewTop >= top && scaledViewTop <= bottom)
                        {
                            region.StartRow = rindex;
                            return 0;
                        }

                        if (scaledViewTop < top)
                            return -1;
                        if (scaledViewTop > bottom)
                            return 1;
                        throw new InvalidOperationException(); // this case should not be reached
                    });
                }

                if (scaledViewBottom > contentBottom)
                {
                    region.EndRow = sheet.rows.Count - 1;
                }
                else
                {
                    var index = sheet.rows.Count / 2;

                    ArrayHelper.QuickFind(index, 0, sheet.rows.Count, rindex =>
                    {
                        var row = sheet.rows[rindex];

                        float top = row.Top;
                        float btn = row.Bottom;

                        if (scaledViewBottom >= top && scaledViewBottom <= btn)
                        {
                            region.EndRow = rindex;
                            return 0;
                        }

                        if (scaledViewBottom < top)
                            return -1;
                        if (scaledViewBottom > btn)
                            return 1;
                        throw new InvalidOperationException(); // this case should not be reached
                    });
                }
            }

            if (viewport.Width > 0 && sheet.cols.Count > 0)
            {
                float contentRight = sheet.cols.Last().Right;

                if (scaledViewLeft > contentRight)
                {
                    region.StartCol = sheet.cols.Count - 1;
                }
                else
                {
                    var index = sheet.cols.Count / 2;

                    ArrayHelper.QuickFind(index, 0, sheet.cols.Count, cindex =>
                    {
                        var col = sheet.cols[cindex];

                        float left = col.Left;
                        float rgt = col.Right;

                        if (scaledViewLeft >= left && scaledViewLeft <= rgt)
                        {
                            region.StartCol = cindex;
                            return 0;
                        }

                        if (scaledViewLeft < left)
                            return -1;
                        if (scaledViewLeft > rgt)
                            return 1;
                        throw new InvalidOperationException(); // this case should not be reached
                    });
                }

                if (scaledViewRight > contentRight)
                {
                    region.EndCol = sheet.cols.Count - 1;
                }
                else
                {
                    var index = sheet.cols.Count / 2;

                    ArrayHelper.QuickFind(index, 0, sheet.cols.Count, cindex =>
                    {
                        var col = sheet.cols[cindex];

                        float left = col.Left;
                        float rgt = col.Right;

                        if (scaledViewRight >= left && scaledViewRight <= rgt)
                        {
                            region.EndCol = cindex;
                            return 0;
                        }

                        if (scaledViewRight < left)
                            return -1;
                        if (scaledViewRight > rgt)
                            return 1;
                        throw new InvalidOperationException(); // this case should not reach
                    });
                }
            }


#if DEBUG
            Debug.Assert(region.EndRow >= region.StartRow);
            Debug.Assert(region.EndCol >= region.StartCol);

            watch.Stop();

            // for unsual visible region
            // when over 200 rows or columns were setted as visible region,
            // we need check for whether the algorithm above has any mistake.
            if (region.Rows > 200 || region.Cols > 200)
                Debug.WriteLine("unusual visible region detected: [row: {0} - {1}, col: {2} - {3}]: {4} ms.",
                    region.StartRow, region.EndRow, region.StartCol, region.EndCol, watch.ElapsedMilliseconds);

            if (watch.ElapsedMilliseconds > 15)
                Debug.WriteLine("update visible region takes " + watch.ElapsedMilliseconds + " ms.");
#endif // DEBUG

            return region;
        }

        #endregion // Visible Region Update

        public override void UpdateController()
        {
#if DEBUG
            var sw = Stopwatch.StartNew();
#endif
            var colOutlineVisible = IsViewVisible(ViewTypes.ColOutline);
            var rowOutlineVisible = IsViewVisible(ViewTypes.RowOutline);

            var colHeadVisible = IsViewVisible(ViewTypes.ColumnHeader);
            var rowHeadVisible = IsViewVisible(ViewTypes.RowHeader);
            var leadHeadVisible = IsViewVisible(ViewTypes.LeadHeader);

            var freezePos = Worksheet.FreezePos;
            var isFrozen = freezePos.Row > 0 || freezePos.Col > 0;

            var outlineSpaceRect = new Rectangle(Bounds.X, Bounds.Y, 0, 0);
            var leadHeadRect = new Rectangle(0, 0, 0, 0);
            var contentRect = new Rectangle(0, 0, 0, 0);

            var scale = ScaleFactor;

            if (colHeadVisible) leadHeadRect.Height = (int)Math.Round(Worksheet.ColHeaderHeight * scale);

            if (rowHeadVisible) leadHeadRect.Width = (int)Math.Round(Worksheet.rowHeaderWidth * scale);

#if OUTLINE
			RGFloat minOutlineButtonScale = (Worksheet.OutlineButtonSize + 3) * Math.Min(scale, 1f);

			if (colOutlineVisible)
			{
				var outlines = Worksheet.outlines[RowOrColumn.Column];

				if (outlines != null)
				{
					// 1
					outlineSpaceRect.Height = (int)Math.Round(outlines.Count * minOutlineButtonScale);
				}
				else
				{
					colOutlineVisible = false;
				}
			}

			if (rowOutlineVisible)
			{
				var outlines = Worksheet.outlines[RowOrColumn.Row];

				if (outlines != null)
				{
					// 2
					outlineSpaceRect.Width = (int)Math.Round(outlines.Count * minOutlineButtonScale);
				}
				else
				{
					rowOutlineVisible = false;
				}
			}
#endif // OUTLINE

            leadHeadRect.X = view.Left + outlineSpaceRect.Width;
            leadHeadRect.Y = view.Top + outlineSpaceRect.Height;
            if (leadHeadVisible) _leadHeadPart.Bounds = leadHeadRect;

            // cells display range boundary
            var contentWidth = view.Right - leadHeadRect.Right;
            var contentHeight = view.Bottom - leadHeadRect.Bottom;

            if (contentWidth < 0) contentWidth = 0;
            if (contentHeight < 0) contentHeight = 0;

            contentRect = new Rectangle(leadHeadRect.Right, leadHeadRect.Bottom, contentWidth, contentHeight);

#if OUTLINE
			if (colOutlineVisible)
			{
				if (rowheadVisible)
				{
					this.colOutlineHeadPart.Bounds = new Rectangle(leadheadRect.X, outlineSpaceRect.Y,
						leadheadRect.Width, outlineSpaceRect.Height);
				}

				this.colOutlinePart2.Bounds = new Rectangle(leadheadRect.Right, outlineSpaceRect.Y,
					contentRect.Width, outlineSpaceRect.Height);
			}

			if (rowOutlineVisible)
			{
				if (colheadVisible)
				{
					this.rowOutlineHeadPart.Bounds = new Rectangle(outlineSpaceRect.X,
						outlineSpaceRect.Bottom, outlineSpaceRect.Width, leadheadRect.Height);
				}
			}
#endif // OUTLINE

            var rightBottomRect = new Rectangle(0, 0, 0, 0);

            if (isFrozen)
            {
                #region Forzen Bounds Layout

                var center = new Point(0, 0);
                var leftTopRect = new Rectangle(0, 0, 0, 0);

                var freezeArea = Worksheet.FreezeArea;
                Rectangle freezeBounds;

                switch (freezeArea)
                {
                    default:
                    case FreezeArea.LeftTop:
                        var gridLoc = GetGridScaleBounds(freezePos);
                        center = new Point(contentRect.X + gridLoc.X, contentRect.Y + gridLoc.Y);
                        break;

                    case FreezeArea.RightBottom:
                        freezeBounds = GetRangeScaledBounds(new GridRegion(
                            freezePos.Row, freezePos.Col,
                            Worksheet.RowCount - 1, Worksheet.ColumnCount - 1));

                        center = new Point(contentRect.Right - freezeBounds.Width,
                            contentRect.Bottom - freezeBounds.Height);
                        break;

                    case FreezeArea.LeftBottom:
                        freezeBounds = GetRangeScaledBounds(new GridRegion(
                            freezePos.Row, freezePos.Col,
                            Worksheet.RowCount - 1, Worksheet.ColumnCount - 1));

                        center = new Point(contentRect.X + freezeBounds.X, contentRect.Bottom - freezeBounds.Height);
                        break;

                    case FreezeArea.RightTop:
                        freezeBounds = GetRangeScaledBounds(new GridRegion(
                            freezePos.Row, freezePos.Col,
                            Worksheet.RowCount - 1, Worksheet.ColumnCount - 1));

                        center = new Point(contentRect.Right - freezeBounds.Width, contentRect.Y + freezeBounds.Y);
                        break;
                }

                if (center.X < contentRect.X) center.X = contentRect.X;
                if (center.Y < contentRect.Y) center.Y = contentRect.Y;
                if (center.X > contentRect.Right) center.X = contentRect.Right;
                if (center.Y > contentRect.Bottom) center.Y = contentRect.Bottom;

                // set left top
                leftTopRect = new Rectangle(contentRect.X, contentRect.Y, center.X - contentRect.X,
                    center.Y - contentRect.Y);

                // set right bottom
                rightBottomRect = new Rectangle(leftTopRect.Right, leftTopRect.Bottom,
                    contentRect.Width - leftTopRect.Width, contentRect.Height - leftTopRect.Height);

                // viewports
                _topLeftViewport.Bounds = leftTopRect;
                _leftBottomViewport.Bounds = new Rectangle(leftTopRect.X, rightBottomRect.Y, leftTopRect.Width,
                    rightBottomRect.Height);
                _rightTopViewport.Bounds = new Rectangle(rightBottomRect.X, leftTopRect.Y, rightBottomRect.Width,
                    leftTopRect.Height);

                // column header
                _colHeaderPart1.Bounds = new Rectangle(_topLeftViewport.Left, _leadHeadPart.Top,
                    _topLeftViewport.Width, Worksheet.ColHeaderHeight * scale);

                // row header
                _rowHeaderPart1.Bounds = new Rectangle(_leadHeadPart.Left, _topLeftViewport.Top,
                    Worksheet.rowHeaderWidth * scale, _topLeftViewport.Height);

#if OUTLINE
				CreateOutlineHeaderViewIfNotExist();

				// column outline
				if (colOutlineVisible)
				{
					this.colOutlinePart1.Bounds = new Rectangle(leftTopRect.X, outlineSpaceRect.Y,
						leftTopRect.Width, outlineSpaceRect.Height);
				}

				// row outline
				if (rowOutlineVisible)
				{
					this.rowOutlinePart1.Bounds = new Rectangle(outlineSpaceRect.X, leftTopRect.Y,
						outlineSpaceRect.Width, leftTopRect.Height);
				}
#endif // OUTLINE

                #endregion // Forzen Bounds Layout
            }
            else
            {
                rightBottomRect = contentRect;
            }

            _rightBottomViewport.Bounds = rightBottomRect;

            _colHeaderPart2.Bounds = new Rectangle(rightBottomRect.X, leadHeadRect.Y, rightBottomRect.Width,
                leadHeadRect.Height);
            _rowHeaderPart2.Bounds = new Rectangle(leadHeadRect.X, rightBottomRect.Y, leadHeadRect.Width,
                rightBottomRect.Height);

#if OUTLINE
			if (rowOutlineVisible)
			{
				this.rowOutlinePart2.Bounds =
 new Rectangle(outlineSpaceRect.X, rightBottomRect.Y, outlineSpaceRect.Width, rightBottomRect.Height);
			}

			if (colOutlineVisible)
			{
				this.colOutlinePart2.Bounds =
 new Rectangle(rightBottomRect.X, outlineSpaceRect.Y, rightBottomRect.Width, outlineSpaceRect.Height);
			}

			if (rowOutlineVisible && colOutlineVisible)
			{
				outlineLeftTopSpace.Bounds = outlineSpaceRect;
			}
#endif // OUTLINE

            //if (this.mainViewport.Width < 0) this.mainViewport.Width = 0;
            //if (this.mainViewport.Height < 0) this.mainViewport.Height = 0;

#if WINFORM || ANDROID
			this.Worksheet.controlAdapter.ScrollBarHorizontalLargeChange = this.scrollHorLarge =
 (int)Math.Round(this.view.Width);
			this.Worksheet.controlAdapter.ScrollBarVerticalLargeChange = this.scrollVerLarge =
 (int)Math.Round(this.view.Height);
#elif WPF
            Worksheet.controlAdapter.ScrollBarHorizontalLargeChange = scrollHorLarge = view.Width;
            Worksheet.controlAdapter.ScrollBarVerticalLargeChange = scrollVerLarge = view.Height;
#endif // WPF

            UpdateVisibleRegion();
            UpdateScrollBarSize();

            // synchronize scale factor
            if (View?.Children != null)
                foreach (var child in View.Children)
                    child.ScaleFactor = View.ScaleFactor;

            view.UpdateView();

            Worksheet.RequestInvalidate();

#if DEBUG
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            if (ms > 0) Debug.WriteLine("update viewport bounds done: " + ms + " ms.");
#endif
        }

        public override void Reset()
        {
            var viewport = view as IViewport;

            if (viewport != null)
            {
                viewport.ViewStart = new Point(0, 0);
                viewport.ScrollX = 0;
                viewport.ScrollY = 0;
            }

            //this.view.ScaleFactor = 1f;
            view.UpdateView();

            _mainVisibleRegion = GridRegion.Empty;
            _frozenVisibleRegion = GridRegion.Empty;

            var adapter = Worksheet.controlAdapter;
            adapter.ScrollBarHorizontalMinimum = 0;
            adapter.ScrollBarVerticalMinimum = 0;
            adapter.ScrollBarHorizontalValue = 0;
            adapter.ScrollBarVerticalValue = 0;
        }

        #endregion // Update

        #region Scroll

        public void HorizontalScroll(double value)
        {
            //if (this.mainViewport.ViewLeft != value)
            //{
            ScrollViews(ScrollDirection.Horizontal, value, -1);
            //}
        }

        public void VerticalScroll(double value)
        {
            //if (this.mainViewport.ViewTop != value)
            //{
            ScrollViews(ScrollDirection.Vertical, -1, value);
            //}
        }

        private double _scrollHorMin;
        private double _scrollHorMax;
        private double _scrollHorLarge;
        private double _scrollHorValue;

        private double _scrollVerMin;
        private double _scrollVerMax;
        private double _scrollVerLarge;
        private double _scrollVerValue;

        public virtual void ScrollOffsetViews(ScrollDirection dir, double offsetX, double offsetY)
        {
            ScrollViews(dir,
                (dir & ScrollDirection.Horizontal) == ScrollDirection.Horizontal ? _scrollHorValue + offsetX : -1,
                (dir & ScrollDirection.Vertical) == ScrollDirection.Vertical ? _scrollVerValue + offsetY : -1);

            SynchronizeScrollBar();
        }

        public virtual void ScrollViews(ScrollDirection dir, double x, double y)
        {
            //if (x == 0 && y == 0) return;
            if (double.IsNaN(x) || double.IsNaN(y)) return;

            if ((dir & ScrollDirection.Horizontal) != ScrollDirection.Horizontal) x = _scrollHorValue;
            if ((dir & ScrollDirection.Vertical) != ScrollDirection.Vertical) y = _scrollVerValue;

            if (x == MainViewport.ScrollX && y == MainViewport.ScrollY) return;

            if (x > _scrollHorMax) x = _scrollHorMax;
            else if (x < 0) x = 0;

            if (y > _scrollVerMax) y = _scrollVerMax;
            else if (y < 0) y = 0;

            // if Control is in edit mode, it is necessary to finish the edit mode
            if (Worksheet.IsEditing) Worksheet.EndEdit(EndEditReason.NormalFinish);

            foreach (var v in view.Children)
            {
                var vp = v as IViewport;
                if (vp != null) vp.ScrollTo(x, y);
            }

            // TODO: Performance Optimization: update visible region by offset 
            UpdateVisibleRegion();

            view.UpdateView();

            Worksheet.RequestInvalidate();

            _scrollHorValue = MainViewport.ScrollX;
            _scrollVerValue = MainViewport.ScrollY;

            Worksheet.workbook?.RaiseWorksheetScrolledEvent(Worksheet, x, y);
        }

        public virtual void ScrollToRange(RangePosition range, CellPosition basePos)
        {
            var view1 = FocusView as Viewport;
            if (view1 != null)
            {
                var rect = Worksheet.GetScaledRangeBounds(range);

                //var rect = this.Worksheet.GetGridBounds(basePos.Row, basePos.Col);
                //rect.Width /= this.Worksheet.scaleFactor;
                //rect.Height /= this.Worksheet.scaleFactor;

                var scale = ScaleFactor;

                var top = view1.ScrollViewTop * scale;
                var bottom = view1.ScrollViewTop * scale + view1.Height;
                var left = view1.ScrollViewLeft * scale;
                var right = view1.ScrollViewLeft * scale + view1.Width;

                double offsetX = 0, offsetY = 0;

                if (rect.Height < view1.Height
                    && (view1.ScrollableDirections & ScrollDirection.Vertical) == ScrollDirection.Vertical)
                    // skip to scroll y if entire row is selected
                    if (range.Rows < Worksheet.rows.Count)
                    {
                        if (rect.Y < top /* && (range.Row <= view.VisibleRegion.startRow)*/)
                            offsetY = (rect.Y - top) / ScaleFactor;
                        else if (rect.Bottom >= bottom /* && (range.EndRow >= view.VisibleRegion.endRow)*/
                                )
                            offsetY = (rect.Bottom - bottom) / ScaleFactor + 1;
                    }

                if (rect.Width < view1.Width
                    && (view1.ScrollableDirections & ScrollDirection.Horizontal) == ScrollDirection.Horizontal)
                    // skip to scroll x if entire column is selected
                    if (range.Cols < Worksheet.cols.Count)
                    {
                        if (rect.X < left /*&& (range.Col <= view.VisibleRegion.startCol)*/)
                            offsetX = (rect.X - left) / ScaleFactor;
                        else if (rect.Right >= right /* && (range.EndCol >= view.VisibleRegion.endCol)*/
                                )
                            offsetX = (rect.Right - right) / ScaleFactor + 1;
                    }

                if (offsetX != 0 || offsetY != 0)
                    ScrollOffsetViews(ScrollDirection.Both, Math.Round(offsetX), Math.Round(offsetY));
            }
        }

        private void UpdateScrollBarSize()
        {
            var scale = ScaleFactor;
            double width = 0, height = 0;

            if (Worksheet.cols.Count > 0)
            {
                width = Worksheet.cols[Worksheet.cols.Count - 1].Right * scale + MainViewport.Left;
#if WPF
                width -= scrollHorLarge;
#endif // WPF 
            }

            if (Worksheet.rows.Count > 0)
            {
                height = Worksheet.rows[Worksheet.rows.Count - 1].Bottom * scale + MainViewport.Top;

                //if (currentFreezePos != CellPosition.Zero)
                //{
                //	height -= this.topLeftViewport.Height;
                //}
#if WPF
                height -= scrollVerLarge;
#endif // WPF
            }

            var maxHorizontal = Math.Max(0, (int)Math.Ceiling(width)) + 1;
            var maxVertical = Math.Max(0, (int)Math.Ceiling(height)) + 1;

            //#if WINFORM || ANDROID
            //			int offHor = maxHorizontal - this.scrollHorMax;
            //			int offVer = maxVertical - this.scrollVerMax;
            //#elif WPF
            //			int offHor = (int)Math.Round(maxHorizontal - this.scrollHorMax);
            //			int offVer = (int)Math.Round(maxVertical - this.scrollVerMax);
            //#elif ANDROID || iOS
            //			RGFloat offHor = maxHorizontal - this.scrollHorMax;
            //			RGFloat offVer = maxVertical - this.scrollVerMax;

            //#endif // WPF
            //if (offHor > 0) offHor = 0;
            //if (offVer > 0) offVer = 0;

            //if (offHor < 0 || offVer < 0)
            //{
            //	ScrollViews(ScrollDirection.Both, offHor, offVer);
            //}

            _scrollHorMax = maxHorizontal;
            _scrollVerMax = maxVertical;

            SynchronizeScrollBar();
        }

        public void SynchronizeScrollBar()
        {
            if (Worksheet == null || Worksheet.controlAdapter == null) return;

            if (_scrollHorValue < _scrollHorMin)
                _scrollHorValue = _scrollHorMin;
            else if (_scrollHorValue > _scrollHorMax) _scrollHorValue = _scrollHorMax;

            if (_scrollVerValue < _scrollVerMin)
                _scrollVerValue = _scrollVerMin;
            else if (_scrollVerValue > _scrollVerMax) _scrollVerValue = _scrollVerMax;

            Worksheet.controlAdapter.ScrollBarHorizontalMaximum = _scrollHorMax;
            Worksheet.controlAdapter.ScrollBarVerticalMaximum = _scrollVerMax;

            Worksheet.controlAdapter.ScrollBarHorizontalMinimum = _scrollHorMin;
            Worksheet.controlAdapter.ScrollBarVerticalMinimum = _scrollVerMin;

            Worksheet.controlAdapter.ScrollBarHorizontalLargeChange = _scrollHorLarge;
            Worksheet.controlAdapter.ScrollBarVerticalLargeChange = _scrollVerLarge;

            Worksheet.controlAdapter.ScrollBarHorizontalValue = _scrollHorValue;
            Worksheet.controlAdapter.ScrollBarVerticalValue = _scrollVerValue;
        }

        #endregion // Scroll

        #region Freeze

        private CellPosition _currentFreezePos = new CellPosition(0, 0);
        private FreezeArea _currentFrozenArea = FreezeArea.None;
        private Point _lastFrozenViewStart;

        public void Freeze(CellPosition freezePos, FreezeArea area = FreezeArea.LeftTop)
        {
            // origin freeze-viewports 
            var cellPosition = Worksheet.GetCellBounds(freezePos);

            if (freezePos == CellPosition.Zero)
            {
                #region Don't freeze

                // restore main-viewport to right-bottom-viewport
                MainViewport = _rightBottomViewport;

                // restore right-bottom viewport settings
                _rightBottomViewport.ViewStart = new Point(0, 0);

                // restore headers viewport settings
                _rowHeaderPart2.ScrollableDirections = ScrollDirection.Vertical;
                _colHeaderPart2.ScrollableDirections = ScrollDirection.Horizontal;

                // hide freeze viewports
                if (_topLeftViewport != null)
                    _topLeftViewport.Visible =
                        _leftBottomViewport.Visible =
                            _rightTopViewport.Visible =
                                false;

                if (_rowHeaderPart1 != null) _rowHeaderPart1.Visible = false;
                if (_colHeaderPart1 != null) _colHeaderPart1.Visible = false;
#if OUTLINE
				if (rowOutlinePart1 != null) rowOutlinePart1.Visible = false;
				if (colOutlinePart1 != null) colOutlinePart1.Visible = false;
#endif // OUTLINE

                #endregion // Don't freeze
            }
            else
            {
                #region Do freeze

                #region Create viewports if not existed

                // right-top cells-viewport
                if (_rightTopViewport == null) AddView(_rightTopViewport = new SheetViewport(this));

                // left-bottom cells-viewport
                if (_leftBottomViewport == null) AddView(_leftBottomViewport = new SheetViewport(this));

                // headers (use InsertPart instead of AddPart to decide the z-orders of viewparts)
                if (_colHeaderPart1 == null) AddView(_colHeaderPart1 = new ColumnHeaderView(this));
                if (_rowHeaderPart1 == null) AddView(_rowHeaderPart1 = new RowHeaderView(this));

                CreateOutlineHeaderViewIfNotExist();

                // left-top cells-viewport
                if (_topLeftViewport == null) AddView(_topLeftViewport = new SheetViewport(this));

                #endregion // Create viewports if not existed

                #region Set viewports view start postion

                _topLeftViewport.ViewStart = new Point(0, 0);
                _leftBottomViewport.ViewStart = new Point(0, cellPosition.Y);
                _rightTopViewport.ViewStart = new Point(cellPosition.X, 0);
                _rightBottomViewport.ViewStart = new Point(cellPosition.X, cellPosition.Y);

                #endregion // Set viewports view start postion

                #region Decides the scroll directions

                // set up viewports by selected freeze position
                switch (Worksheet.FreezeArea)
                {
                    default:
                    case FreezeArea.LeftTop:
                        _topLeftViewport.ScrollableDirections = ScrollDirection.None;
                        _leftBottomViewport.ScrollableDirections = ScrollDirection.Vertical;
                        _rightTopViewport.ScrollableDirections = ScrollDirection.Horizontal;

                        _colHeaderPart1.ScrollableDirections = ScrollDirection.None;
                        _rowHeaderPart1.ScrollableDirections = ScrollDirection.None;
                        _colHeaderPart2.ScrollableDirections = ScrollDirection.Horizontal;
                        _rowHeaderPart2.ScrollableDirections = ScrollDirection.Vertical;

                        MainViewport = _rightBottomViewport;
                        break;

                    case FreezeArea.RightBottom:
                        _leftBottomViewport.ScrollableDirections = ScrollDirection.Horizontal;
                        _rightTopViewport.ScrollableDirections = ScrollDirection.Vertical;
                        _rightBottomViewport.ScrollableDirections = ScrollDirection.None;

                        _colHeaderPart1.ScrollableDirections = ScrollDirection.Horizontal;
                        _rowHeaderPart1.ScrollableDirections = ScrollDirection.Vertical;
                        _colHeaderPart2.ScrollableDirections = ScrollDirection.None;
                        _rowHeaderPart2.ScrollableDirections = ScrollDirection.None;

                        MainViewport = _topLeftViewport;
                        break;

                    case FreezeArea.LeftBottom:
                        _topLeftViewport.ScrollableDirections = ScrollDirection.Vertical;
                        _leftBottomViewport.ScrollableDirections = ScrollDirection.None;
                        _rightBottomViewport.ScrollableDirections = ScrollDirection.Horizontal;

                        _colHeaderPart1.ScrollableDirections = ScrollDirection.None;
                        _rowHeaderPart1.ScrollableDirections = ScrollDirection.Vertical;
                        _colHeaderPart2.ScrollableDirections = ScrollDirection.Horizontal;
                        _rowHeaderPart2.ScrollableDirections = ScrollDirection.None;

                        MainViewport = _rightTopViewport;
                        break;

                    case FreezeArea.RightTop:
                        _topLeftViewport.ScrollableDirections = ScrollDirection.Horizontal;
                        _rightTopViewport.ScrollableDirections = ScrollDirection.None;
                        _rightBottomViewport.ScrollableDirections = ScrollDirection.Vertical;

                        _colHeaderPart1.ScrollableDirections = ScrollDirection.Horizontal;
                        _rowHeaderPart1.ScrollableDirections = ScrollDirection.None;
                        _colHeaderPart2.ScrollableDirections = ScrollDirection.None;
                        _rowHeaderPart2.ScrollableDirections = ScrollDirection.Vertical;

                        MainViewport = _leftBottomViewport;
                        break;
                }

                #endregion // Decides the scroll directions for viewports

                #region Set viewports visibility

                _topLeftViewport.Visible =
                    _leftBottomViewport.Visible =
                        _rightTopViewport.Visible =
                            true;

                _colHeaderPart1.Visible = _colHeaderPart2.Visible;
                _rowHeaderPart1.Visible = _rowHeaderPart2.Visible;

#if OUTLINE
				if (this.colOutlinePart1 != null && this.colOutlinePart2 != null)
				{
					this.colOutlinePart1.Visible = colOutlinePart2.Visible;
				}
				if (this.rowOutlinePart1 != null && this.rowOutlinePart2 != null)
				{
					this.rowOutlinePart1.Visible = rowOutlinePart2.Visible;
				}
#endif // OUTLINE

                #endregion // Set viewports visibility

                #endregion // Do freeze
            }

            // Scrollable direction for main viewport always should be Both
            MainViewport.ScrollableDirections = ScrollDirection.Both;

            #region Synchronize other viewports

            if (_rowHeaderPart1 != null) _rowHeaderPart1.ViewTop = _topLeftViewport.ViewTop;
            if (_colHeaderPart1 != null) _colHeaderPart1.ViewLeft = _topLeftViewport.ViewLeft;

            _colHeaderPart2.ViewLeft = _rightBottomViewport.ViewLeft;
            _rowHeaderPart2.ViewTop = _rightBottomViewport.ViewTop;

            #endregion Synchronize other viewports

            #region Outline

#if OUTLINE
			if (colOutlinePart2 != null)
			{
				colOutlinePart2.ViewLeft = rightBottomViewport.ViewLeft;
				colOutlinePart2.ScrollableDirections = colHeaderPart2.ScrollableDirections;

				if (colOutlinePart1 != null)
				{
					this.colOutlinePart1.ViewStart = this.colHeaderPart1.ViewStart;
					this.colOutlinePart1.ScrollableDirections = this.colHeaderPart1.ScrollableDirections;
				}
			}

			if (rowOutlinePart2 != null)
			{
				rowOutlinePart2.ViewTop = rightBottomViewport.ViewTop;
				rowOutlinePart2.ScrollableDirections = rowHeaderPart2.ScrollableDirections;

				if (rowOutlinePart1 != null)
				{
					this.rowOutlinePart1.ViewStart = this.rowHeaderPart1.ViewStart;
					this.rowOutlinePart1.ScrollableDirections = this.rowHeaderPart1.ScrollableDirections;
				}
			}
#endif // OUTLINE

            #endregion // Outline

            #region Update Scrollbar positions

            // scroll bars start at view-start of the main-viewport
            Worksheet.controlAdapter.ScrollBarHorizontalMinimum =
                0; // this.scrollHorMin = (int)Math.Round(this.mainViewport.ViewLeft);
            Worksheet.controlAdapter.ScrollBarVerticalMinimum =
                0; // this.scrollVerMin = (int)Math.Round(this.mainViewport.ViewTop);

            //int hlc = (int)(this.view.Width - rightBottomViewport.ScrollViewLeft);
            //if (hlc < 0) hlc = 0;
            //int vlc = (int)(this.view.Height - rightBottomViewport.ScrollViewTop);
            //if (vlc < 0) vlc = 0;

            //this.Worksheet.controlAdapter.ScrollBarHorizontalLargeChange = this.scrollHorLarge = hlc;
            //this.Worksheet.controlAdapter.ScrollBarVerticalLargeChange = this.scrollVerLarge = vlc;
            Worksheet.controlAdapter.ScrollBarHorizontalLargeChange = Math.Round(MainViewport.Width);
            Worksheet.controlAdapter.ScrollBarVerticalLargeChange = Math.Round(MainViewport.Height);

            #endregion // Update Scrollbar positions

            _currentFreezePos = freezePos;
            _currentFrozenArea = area;

            UpdateController();
        }

        private void CreateOutlineHeaderViewIfNotExist()
        {
#if OUTLINE
			// outline-views 
			if (rowOutlinePart1 == null && rowOutlinePart2 != null)
			{
				AddView(rowOutlinePart1 = new RowOutlineView(this));

				// move row outline header part to the topmost
				this.RemoveView(this.rowOutlineHeadPart);
				AddView(this.rowOutlineHeadPart);
			}

			if (colOutlinePart1 == null && colOutlinePart2 != null)
			{
				AddView(colOutlinePart1 = new ColumnOutlinePart(this));

				// move column outline header part to the topmost
				this.RemoveView(this.colOutlineHeadPart);
				AddView(this.colOutlineHeadPart);

				// adjust z-index of outline left top corner space
				if (this.outlineLeftTopSpace != null)
				{
					this.RemoveView(this.outlineLeftTopSpace);
					AddView(this.outlineLeftTopSpace);
				}
			}
#endif // OUTLINE
        }

        #endregion // Freeze

        #region Outline

#if OUTLINE
		private RowOutlineView rowOutlinePart2;
		private RowOutlineView rowOutlinePart1;
		private RowOutlineHeaderView rowOutlineHeadPart;
		private ColumnOutlinePart colOutlinePart2;
		private ColumnOutlinePart colOutlinePart1;
		private ColumnOutlineHeadPart colOutlineHeadPart;
		private OutlineLeftTopSpace outlineLeftTopSpace;
#endif // OUTLINE

        #endregion // Outline

        #region Draw

        public override void Draw(CellDrawingContext dc)
        {
            base.Draw(dc);

            var g = dc.Graphics;

            #region Freeze Split Line

            if (view != null
                && Worksheet.HasSettings(WorksheetSettings.View_ShowFrozenLine))
            {
                var freezePos = Worksheet.FreezePos;

                if (freezePos.Col > 0)
                    g.DrawLine(_leftBottomViewport.Right, view.Top, _leftBottomViewport.Right, view.Bottom,
                        SolidColor.Gray);

                if (freezePos.Row > 0)
                    g.DrawLine(view.Left, _rightTopViewport.Bottom, view.Right, _rightTopViewport.Bottom,
                        SolidColor.Gray);
            }

            #endregion // Freeze Split Line
        }

        #endregion // Draw
    }
}