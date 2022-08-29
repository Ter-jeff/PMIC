using System;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;

namespace MyWpf.Converter
{
    [ValueConversion(typeof(Orientation), typeof(Dock))]
    public class OrientationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return Dock.Left;
            var orientation = (Orientation)value;
            if (orientation == Orientation.Horizontal)
                return Dock.Left;
            if (orientation == Orientation.Vertical)
                return Dock.Top;
            return Dock.Left;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
