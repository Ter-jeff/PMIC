using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MyWpf.Controls
{
    public class MyImageButton : Button
    {
        static MyImageButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MyImageButton),
               new FrameworkPropertyMetadata(typeof(MyImageButton)));
        }

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        public static readonly DependencyProperty ImageSourceProperty =
            DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(MyImageButton), new PropertyMetadata(null));
    }
}
