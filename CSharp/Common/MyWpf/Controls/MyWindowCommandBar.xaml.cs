using MyWpf.Controls.Ribbon.UI;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MyWpf.Controls
{
    public partial class MyWindowCommandBar
    {
        public const string SettingBarName = "MySettingBar";

        public UIElement SettingBar
        {
            get { return (UIElement)GetValue(SettingBarProperty); }
            set { SetValue(SettingBarProperty, value); }
        }

        public static readonly DependencyProperty SettingBarProperty =
            DependencyProperty.Register("SettingBar", typeof(UIElement), typeof(MyWindowCommandBar), new PropertyMetadata(null));

        public MyWindowCommandBar()
        {
            InitializeComponent();

            CommandBindings.Add(new CommandBinding(WindowCommands.Maximize, OnMaximize));
            CommandBindings.Add(new CommandBinding(WindowCommands.Minimize, OnMinimize));
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Close, OnClose));
        }

        public ImageSource Icon
        {
            get { return (ImageSource)GetValue(IconProperty); }
            set { SetValue(IconProperty, value); }
        }

        public static readonly DependencyProperty IconProperty =
            DependencyProperty.Register("Icon", typeof(ImageSource), typeof(MyWindowCommandBar), new PropertyMetadata(null));

        public string Title
        {
            get
            {
                return (string)GetValue(TitleProperty);
            }
            set
            {
                SetValue(TitleProperty, value);
            }
        }

        public static readonly DependencyProperty TitleProperty = DependencyProperty.Register("Title", typeof(string), typeof(MyWindowCommandBar), new PropertyMetadata(default(string)));

        public HorizontalAlignment TitleHorizontalAlignment
        {
            get { return (HorizontalAlignment)GetValue(TitleHorizontalAlignmentProperty); }
            set { SetValue(TitleHorizontalAlignmentProperty, value); }
        }

        public T FindParentOfType<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentDepObj = child;
            do
            {
                parentDepObj = VisualTreeHelper.GetParent(parentDepObj);
                T parent = parentDepObj as T;
                if (parent != null) return parent;
            }
            while (parentDepObj != null);
            return null;
        }

        public static readonly DependencyProperty TitleHorizontalAlignmentProperty =
            DependencyProperty.Register("TitleHorizontalAlignment", typeof(HorizontalAlignment), typeof(MyWindowCommandBar), new PropertyMetadata(HorizontalAlignment.Left));

        private void OnClose(object sender, RoutedEventArgs e)
        {
            var myWindow = FindParentOfType<MyWindow>(this);
            if (myWindow != null)
                myWindow.Close();
            e.Handled = true;
        }

        private void OnMinimize(object sender, ExecutedRoutedEventArgs e)
        {
            var myWindow = FindParentOfType<MyWindow>(this);
            if (myWindow != null)
                myWindow.WindowState = WindowState.Minimized;
            e.Handled = true;
        }

        private void OnMaximize(object sender, ExecutedRoutedEventArgs e)
        {
            var myWindow = FindParentOfType<MyWindow>(this);
            if (myWindow != null)
                if (myWindow.WindowState == WindowState.Normal)
                    myWindow.WindowState = WindowState.Maximized;
                else if (myWindow.WindowState == WindowState.Maximized)
                    myWindow.WindowState = WindowState.Normal;
                else
                    myWindow.WindowState = WindowState.Maximized;
            e.Handled = true;
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            if (SettingBar == null)
            {
                var icon = new Image();
                icon.Source = Icon ??
                              new BitmapImage(new Uri("/MyWpf;component/Resources/Teradyne_T.ico", UriKind.Relative));
                icon.Stretch = Stretch.UniformToFill;
                MySettingBar.Children.Add(icon);
            }
            else
            {
                MySettingBar.Children.Add(SettingBar);
            }
        }
    }
}
