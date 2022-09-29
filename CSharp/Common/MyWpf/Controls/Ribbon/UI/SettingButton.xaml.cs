using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MyWpf.Controls.Ribbon.UI
{
    public class SettingButton : Button
    {
        static SettingButton()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SettingButton),
               new FrameworkPropertyMetadata(typeof(SettingButton)));
        }

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        public static readonly DependencyProperty ImageSourceProperty =
            DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(SettingButton), new PropertyMetadata(null));
    }
}
