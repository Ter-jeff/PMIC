using System;
using System.Windows;
using System.Windows.Input;

namespace MyWpf.Controls
{
    public class MyWindow : MyWindowBase
    {
        public static readonly DependencyProperty TitleHeightProperty = DependencyProperty.Register("TitleHeight",
            typeof(int), typeof(MyWindow), new PropertyMetadata(default(int)));

        static MyWindow()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MyWindow),
                new FrameworkPropertyMetadata(typeof(MyWindow)));
        }

        public int TitleHeight
        {
            get { return (int)GetValue(TitleHeightProperty); }
            set { SetValue(TitleHeightProperty, value); }
        }

        public override void OnApplyTemplate()
        {
            var myWindowCommandBar = GetTemplateChild("MyWindowCommandBar") as MyWindowCommandBar;
            if (myWindowCommandBar != null)
                myWindowCommandBar.HelpButton.Click += Help_Click;
            if (myWindowCommandBar != null)
                myWindowCommandBar.MouseDoubleClick += Window_MouseDoubleClick;

        }

        public event EventHandler<EventArgs> HelpButtonClick;

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            if (HelpButtonClick != null)
            {
                HelpButtonClick(this, EventArgs.Empty);
            }
        }

        private void Window_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }
    }
}
