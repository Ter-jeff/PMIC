using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace MyWpf.Controls.Ribbon.UI
{
    public class RibbonButton : ToggleButton
    {
        static RibbonButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(RibbonButton),
               new FrameworkPropertyMetadata(typeof(RibbonButton)));
        }

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        public static readonly DependencyProperty ImageSourceProperty =
            DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(RibbonButton), new PropertyMetadata(null));

        public string Header
        {
            get { return (string)GetValue(HeaderProperty); }
            set { SetValue(HeaderProperty, value); }
        }

        public static readonly DependencyProperty HeaderProperty =
            DependencyProperty.Register("Header", typeof(string), typeof(RibbonButton), new PropertyMetadata(null));
    }
}
