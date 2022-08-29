using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;
using SpreedSheet.Control;

namespace SpreedSheet.WPF
{
    internal class SheetTabControl : Grid, ISheetTabControl
    {
        private readonly Image _newSheetImage;
        private readonly DispatcherTimer _scrollTimer;
        private bool _scrollLeftDown;
        private bool _scrollRightDown;

        private bool _splitterMoving;

        internal Grid Canvas = new Grid
        {
            Width = 0,
            VerticalAlignment = VerticalAlignment.Top,
            HorizontalAlignment = HorizontalAlignment.Left
        };

        public SheetTabControl()
        {
            Background = SystemColors.ControlBrush;
            BorderColor = Colors.DeepSkyBlue;

            ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            ColumnDefinitions.Add(new ColumnDefinition());
            ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
            ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(5) });

            Border left = new ArrowBorder(this)
            {
                Child = new Polygon
                {
                    Points = new PointCollection(
                        new[]
                        {
                            new Point(6, 0),
                            new Point(0, 5),
                            new Point(6, 10)
                        }),
                    Fill = SystemColors.ControlDarkDarkBrush,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(4, 0, 0, 0)
                },
                Background = SystemColors.ControlBrush
            };

            Border right = new ArrowBorder(this)
            {
                Child = new Polygon
                {
                    Points = new PointCollection(
                        new[]
                        {
                            new Point(0, 0),
                            new Point(6, 5),
                            new Point(0, 10)
                        }),
                    Fill = SystemColors.ControlDarkDarkBrush,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0)
                },
                Background = SystemColors.ControlBrush
            };

            Canvas.RenderTransform = new TranslateTransform(0, 0);
            Children.Add(Canvas);
            SetColumn(Canvas, 2);

            Children.Add(left);
            SetColumn(left, 0);

            Children.Add(right);
            SetColumn(right, 1);

            var imageSource = new BitmapImage();
            var imageHoverSource = new BitmapImage();

            imageSource.BeginInit();
            imageSource.StreamSource = new MemoryStream(Properties.Resources.NewBuildDefinition_8952_inactive_png);
            imageSource.EndInit();

            imageHoverSource.BeginInit();
            imageHoverSource.StreamSource = new MemoryStream(Properties.Resources.NewBuildDefinition_8952_png);
            imageHoverSource.EndInit();

            _newSheetImage = new Image
            {
                Source = imageSource,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2),
                Cursor = Cursors.Hand
            };

            _newSheetImage.MouseEnter += (s, e) => _newSheetImage.Source = imageHoverSource;
            _newSheetImage.MouseLeave += (s, e) => _newSheetImage.Source = imageSource;
            _newSheetImage.MouseDown += (s, e) =>
            {
                if (NewSheetClick != null) NewSheetClick(this, null);
            };

            Children.Add(_newSheetImage);
            SetColumn(_newSheetImage, 3);

            var rightThumb = new Border
            {
                Child = new RightThumb(this),
                Cursor = Cursors.SizeWE,
                Background = SystemColors.ControlBrush,
                Margin = new Thickness(0, 1, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            Children.Add(rightThumb);
            SetColumn(rightThumb, 4);

            _scrollTimer = new DispatcherTimer
            {
                Interval = new TimeSpan(0, 0, 0, 0, 10)
            };

            _scrollTimer.Tick += (s, e) =>
            {
                var tt = Canvas.Margin.Left;

                if (_scrollLeftDown)
                {
                    if (tt < 0)
                    {
                        tt += 5;
                        if (tt > 0) tt = 0;
                    }
                }
                else if (_scrollRightDown)
                {
                    var max = ColumnDefinitions[2].ActualWidth - Canvas.Width;

                    if (tt > max)
                    {
                        tt -= 5;
                        if (tt < max) tt = max;
                    }
                }

                if (Canvas.Margin.Left != tt) Canvas.Margin = new Thickness(tt, 0, 0, 0);
            };

            left.MouseDown += (s, e) =>
            {
                _scrollRightDown = false;
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    _scrollLeftDown = true;
                    _scrollTimer.IsEnabled = true;
                }
                else if (e.RightButton == MouseButtonState.Pressed)
                {
                    if (SheetListClick != null) SheetListClick(this, null);
                }
            };
            left.MouseUp += (s, e) =>
            {
                _scrollTimer.IsEnabled = false;
                _scrollLeftDown = false;
            };

            right.MouseDown += (s, e) =>
            {
                _scrollLeftDown = false;
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    _scrollRightDown = true;
                    _scrollTimer.IsEnabled = true;
                }
                else if (e.RightButton == MouseButtonState.Pressed)
                {
                    if (SheetListClick != null) SheetListClick(this, null);
                }
            };
            right.MouseUp += (s, e) =>
            {
                _scrollTimer.IsEnabled = false;
                _scrollRightDown = false;
            };

            rightThumb.MouseDown += (s, e) =>
            {
                _splitterMoving = true;
                rightThumb.CaptureMouse();
            };
            rightThumb.MouseMove += (s, e) =>
            {
                if (_splitterMoving)
                    if (SplitterMoving != null)
                        SplitterMoving(this, null);
            };
            rightThumb.MouseUp += (s, e) =>
            {
                _splitterMoving = false;
                rightThumb.ReleaseMouseCapture();
            };
        }

        /// <summary>
        ///     Determine whether or not allow to move tab by dragging mouse
        /// </summary>
        public bool AllowDragToMove { get; set; }

        public void ScrollToItem(int index)
        {
            // TODO!

            //double width = this.ColumnDefinitions[2].ActualWidth;
            //int visibleWidth = this.ClientRectangle.Width - leftPadding - rightPadding;

            //if (rect.Width > visibleWidth || rect.Left < this.viewScroll + leftPadding)
            //{
            //	this.viewScroll = rect.Left - leftPadding;
            //}
            //else if (rect.Right - this.viewScroll > this.ClientRectangle.Right - rightPadding)
            //{
            //	this.viewScroll = rect.Right - this.ClientRectangle.Width + leftPadding;
            //}
        }

        public double ControlWidth { get; set; }

        public event EventHandler<SheetTabMovedEventArgs> TabMoved;

        public event EventHandler SelectedIndexChanged;

        public event EventHandler SplitterMoving;

        public event EventHandler SheetListClick;

        public event EventHandler NewSheetClick;

        public event EventHandler<SheetTabMouseEventArgs> TabMouseDown;

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            Clip = new RectangleGeometry(new Rect(0, 0, RenderSize.Width, RenderSize.Height));

            Canvas.Height = Height - 2;
        }

        protected override void OnInitialized(EventArgs e)
        {
            base.OnInitialized(e);
        }

        public double TranslateScrollPoint(int p)
        {
            return Canvas.RenderTransform.Transform(new Point(p, 0)).X;
        }

        public Rect GetItemBounds(int index)
        {
            if (index < 0 || index > Canvas.Children.Count - 1) throw new ArgumentOutOfRangeException("index");

            var tab = Canvas.Children[index];

            return new Rect(tab.PointToScreen(new Point(0, 0)), RenderSize);
        }

        public void MoveItem(int index, int targetIndex)
        {
            if (index < 0 || index > Canvas.Children.Count - 1) throw new ArgumentOutOfRangeException("index");

            var tab = Canvas.Children[index];

            Canvas.Children.RemoveAt(index);

            if (targetIndex > index) targetIndex--;

            Canvas.Children.Insert(targetIndex, tab);
        }

        #region Dependency Properties

        public static readonly DependencyProperty SelectedBackColorProperty =
            DependencyProperty.Register("SelectedBackColor", typeof(Color), typeof(SheetTabControl));

        public Color SelectedBackColor
        {
            get { return (Color)GetValue(SelectedBackColorProperty); }
            set { SetValue(SelectedBackColorProperty, value); }
        }

        public static readonly DependencyProperty SelectedTextColorProperty =
            DependencyProperty.Register("SelectedTextColor", typeof(Color), typeof(SheetTabControl));

        public Color SelectedTextColor
        {
            get { return (Color)GetValue(SelectedTextColorProperty); }
            set { SetValue(SelectedTextColorProperty, value); }
        }

        public static readonly DependencyProperty BorderColorProperty =
            DependencyProperty.Register("BorderColor", typeof(Color), typeof(SheetTabControl));

        public Color BorderColor
        {
            get { return (Color)GetValue(BorderColorProperty); }
            set { SetValue(BorderColorProperty, value); }
        }

        public static readonly DependencyProperty SelectedIndexProperty =
            DependencyProperty.Register("SelectedIndex", typeof(int), typeof(SheetTabControl));

        public int SelectedIndex
        {
            get { return (int)GetValue(SelectedIndexProperty); }

            set
            {
                var tabContainer = Canvas;

                var currentIndex = SelectedIndex;

                if (currentIndex >= 0 && currentIndex < tabContainer.Children.Count)
                {
                    var tab = tabContainer.Children[currentIndex] as SheetTabItem;
                    if (tab != null) tab.IsSelected = false;
                }

                SetValue(SelectedIndexProperty, value);
                currentIndex = value;

                if (currentIndex >= 0 && currentIndex < tabContainer.Children.Count)
                {
                    var tab = tabContainer.Children[currentIndex] as SheetTabItem;
                    if (tab != null) tab.IsSelected = true;
                }

                if (SelectedIndexChanged != null) SelectedIndexChanged(this, null);
            }
        }

        public bool NewButtonVisible
        {
            get { return _newSheetImage.Visibility == Visibility.Visible; }
            set { _newSheetImage.Visibility = value ? Visibility.Visible : Visibility.Hidden; }
        }

        #endregion // Dependency Properties

        #region Tab Management

        public void AddTab(string title)
        {
            var index = Canvas.Children.Count;
            InsertTab(index, title);
        }

        public void InsertTab(int index, string title)
        {
            var tab = new SheetTabItem(this, title)
            {
                Height = Canvas.Height
            };

            Canvas.Width += tab.Width + 1;
            Canvas.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            //Canvas.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(tab.Width + 1) });
            for (var i = 0; i < Canvas.Children.Count; i++)
                if (i >= index)
                {
                    SetColumn(Canvas.Children[i], i + 1);
                    ((SheetTabItem)Canvas.Children[i]).IsSelected = false;
                }

            Canvas.Children.Insert(index, tab);

            SetColumn(tab, index);

            tab.MouseDown += Tab_MouseDown;

            //if (Canvas.Children.Count == 1)
            //{
            tab.IsSelected = true;
            //}
        }

        private void Tab_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var index = Canvas.Children.IndexOf((UIElement)sender);

            var arg = new SheetTabMouseEventArgs
            {
                Handled = false,
                Location = e.GetPosition(this),
                Index = index,
                MouseButtons = WpfUtility.ConvertToUiMouseButtons(e)
            };

            if (TabMouseDown != null) TabMouseDown(this, arg);

            if (!arg.Handled) SelectedIndex = index;
        }

        public void RemoveTab(int index)
        {
            var tab = (SheetTabItem)Canvas.Children[index];

            Canvas.Children.RemoveAt(index);
            Canvas.ColumnDefinitions.RemoveAt(index);

            for (var i = index; i < Canvas.Children.Count; i++) SetColumn(Canvas.Children[i], i);

            Canvas.Width -= tab.Width;
        }

        public void UpdateTab(int index, string title, Color backColor, Color textColor)
        {
            var item = Canvas.Children[index] as SheetTabItem;
            if (item != null)
            {
                item.ChangeTitle(title);
                Canvas.ColumnDefinitions[index].Width = GridLength.Auto;
                //Canvas.ColumnDefinitions[index].Width = new GridLength(item.Width + 1);

                item.BackColor = backColor;
                item.TextColor = textColor;
            }
        }

        public void ClearTabs()
        {
            Canvas.Children.Clear();
            Canvas.ColumnDefinitions.Clear();
            Canvas.Width = 0;
        }

        public int TabCount
        {
            get { return Canvas.Children.Count; }
        }

        #endregion // Tab Management

        #region Paint

        private readonly GuidelineSet _gls = new GuidelineSet();

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);

            var g = dc;

            _gls.GuidelinesX.Clear();
            _gls.GuidelinesY.Clear();

            _gls.GuidelinesX.Add(0.5);
            _gls.GuidelinesX.Add(RenderSize.Width + 0.5);
            _gls.GuidelinesY.Add(0.5);
            _gls.GuidelinesY.Add(RenderSize.Height + 0.5);

            g.PushGuidelineSet(_gls);

            var p = new Pen(new SolidColorBrush(BorderColor), 1);

            g.DrawLine(p, new Point(0, 0), new Point(RenderSize.Width, 0));

            g.Pop();
        }

        #endregion // Paint
    }

    internal class SheetTabItem : Decorator
    {
        public static readonly DependencyProperty IsSelectedProperty =
            DependencyProperty.Register("IsSelected", typeof(bool), typeof(SheetTabItem));

        private readonly GuidelineSet _gls = new GuidelineSet();
        private readonly SheetTabControl _owner;

        public SheetTabItem(SheetTabControl owner, string title)
        {
            _owner = owner;

            SnapsToDevicePixels = true;

            ChangeTitle(title);
        }

        public bool IsSelected
        {
            get { return (bool)GetValue(IsSelectedProperty); }
            set
            {
                var currentValue = (bool)GetValue(IsSelectedProperty);

                if (currentValue != value)
                {
                    SetValue(IsSelectedProperty, value);
                    InvalidateVisual();
                }
            }
        }

        public Color BackColor { get; set; }
        public Color TextColor { get; set; }

        public void ChangeTitle(string title)
        {
            var label = new TextBlock
            {
                Text = title,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Background = Brushes.Transparent,
                Margin = new Thickness(10, 0, 10, 0)
            };
            //label.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            Child = label;
            //Width = label.DesiredSize.Width + 9;
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            var g = drawingContext;

            var right = RenderSize.Width;
            var bottom = RenderSize.Height;

            _gls.GuidelinesX.Clear();
            _gls.GuidelinesY.Clear();

            _gls.GuidelinesX.Add(0.5);
            _gls.GuidelinesX.Add(right + 0.5);
            _gls.GuidelinesY.Add(0.5);
            _gls.GuidelinesY.Add(bottom + 0.5);

            g.PushGuidelineSet(_gls);

            Brush b = new SolidColorBrush(_owner.BorderColor);
            var p = new Pen(b, 1);

            if (IsSelected)
            {
                g.DrawRectangle(
                    BackColor.A > 0 ? new SolidColorBrush(BackColor) : Brushes.White,
                    null, new Rect(0, 0, right, bottom));

                g.DrawLine(p, new Point(0, 0), new Point(0, bottom));
                g.DrawLine(p, new Point(right, 0), new Point(right, bottom));

                g.DrawLine(p, new Point(0, bottom), new Point(right, bottom));

                g.DrawLine(new Pen(Brushes.White, 1), new Point(1, 0), new Point(right, 0));
            }
            else
            {
                g.DrawRectangle(
                    BackColor.A > 0 ? new SolidColorBrush(BackColor) : SystemColors.ControlBrush,
                    null, new Rect(0, 0, right, bottom));

                var index = _owner.Canvas.Children.IndexOf(this);

                if (index > 0)
                    g.DrawLine(new Pen(SystemColors.ControlDarkDarkBrush, 1), new Point(0, 2),
                        new Point(0, bottom - 2));

                // top border
                g.DrawLine(p, new Point(0, 0), new Point(right, 0));
            }

            g.Pop();
        }
    }

    internal class ArrowBorder : Border
    {
        private readonly GuidelineSet _gls = new GuidelineSet { GuidelinesY = new DoubleCollection(new[] { 0.5 }) };
        private readonly SheetTabControl _owner;

        public ArrowBorder(SheetTabControl owner)
        {
            _owner = owner;

            SnapsToDevicePixels = true;
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);

            var g = dc;

            g.PushGuidelineSet(_gls);

            g.DrawLine(new Pen(new SolidColorBrush(_owner.BorderColor), 1),
                new Point(0, 0), new Point(RenderSize.Width, 0));

            g.Pop();
        }
    }

    internal class RightThumb : FrameworkElement
    {
        private readonly SheetTabControl _owner;

        public RightThumb(SheetTabControl owner)
        {
            _owner = owner;
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            return new Size(5, 0);
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            var g = drawingContext;

            var b = new SolidColorBrush(_owner.BorderColor);
            var p = new Pen(b, 1);

            for (double y = 3; y < RenderSize.Height - 3; y += 4)
                g.DrawRectangle(SystemColors.ControlDarkBrush, null, new Rect(0, y, 2, 2));

            var right = RenderSize.Width;

            var gls = new GuidelineSet();
            gls.GuidelinesX.Add(right + 0.5);
            g.PushGuidelineSet(gls);

            g.DrawLine(p, new Point(right, 0), new Point(right, RenderSize.Height));

            g.Pop();
        }
    }
}