using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MyWpf.Controls
{
    public class MyMenuButton : RadioButton
    {
        static MyMenuButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MyMenuButton),
               new FrameworkPropertyMetadata(typeof(MyMenuButton)));
        }

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        public static readonly DependencyProperty ImageSourceProperty =
            DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(MyMenuButton), new PropertyMetadata(null));
    }
}
