using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MyWpf.Converters
{
    [ValueConversion(typeof(string), typeof(Visibility))]
    public class TextToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return Visibility.Collapsed;
            var text = (string)value;
            if (string.IsNullOrEmpty(text))
                return Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}