using System.Windows;
using System.Windows.Controls.Primitives;

namespace MyWpf.Controls.Ribbon.UI
{
    public class RibbonToggleButton : ToggleButton
    {
        static RibbonToggleButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(RibbonToggleButton),
               new FrameworkPropertyMetadata(typeof(RibbonToggleButton)));
        }

        public bool Down
        {
            get { return (bool)IsChecked; }
            set { IsChecked = value; }
        }

        public string Header
        {
            get { return (string)GetValue(HeaderProperty); }
            set { SetValue(HeaderProperty, value); }
        }

        public static readonly DependencyProperty HeaderProperty =
            DependencyProperty.Register("Header", typeof(string), typeof(RibbonToggleButton), new PropertyMetadata(null));
    }
}
