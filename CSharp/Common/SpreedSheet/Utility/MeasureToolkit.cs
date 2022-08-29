#define WPF

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WPF

namespace unvell.ReoGrid.Utility
{
    internal static class MeasureToolkit
    {
        // 1 inch = 2.54 cm
        private const double _cm_pre_inch = 2.54f;
        private const double _windows_standard_dpi = 96f;

        public const int _emi_in_inch = 914400;

        public static double InchToPixel(double inch, double dpi)
        {
            return inch * dpi;
        }

        public static double PixelToInch(double px, double dpi)
        {
            return px / dpi;
        }

        public static double InchToPixel(double inch)
        {
            return InchToPixel(inch, _windows_standard_dpi);
        }

        public static double PixelToInch(double px)
        {
            return PixelToInch(px, _windows_standard_dpi);
        }

        public static double InchToCM(double inch)
        {
            return inch * _cm_pre_inch;
        }

        public static double CMToInch(double cm)
        {
            return cm / _cm_pre_inch;
        }

        public static double CMToPixel(double cm, double dpi)
        {
            return cm * dpi / _cm_pre_inch;
        }

        public static double PixelToCM(double px, double dpi)
        {
            return px * _cm_pre_inch / dpi;
        }

        public static double CMTOPixel(double cm)
        {
            return CMToPixel(cm, _windows_standard_dpi);
        }

        public static double PixelToCM(double px)
        {
            return PixelToCM(px, _windows_standard_dpi);
        }

        public static double EMUToPixel(int emu, double dpi)
        {
            return (double)emu / _emi_in_inch * dpi;
        }

        public static int PixelToEMU(double pixel, double dpi)
        {
            return (int)(pixel * _emi_in_inch / dpi);
        }
    }
}