#define WPF


#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;

#endif // WPF

namespace unvell.ReoGrid.Drawing.Text
{
    /// <summary>
    ///     Font style
    /// </summary>
    public enum FontStyles : byte
    {
        /// <summary>
        ///     Regular
        /// </summary>
        Regular = 0,

        /// <summary>
        ///     Bold
        /// </summary>
        Bold = 1,

        /// <summary>
        ///     Italic
        /// </summary>
        Italic = 2,

        /// <summary>
        ///     Underline
        /// </summary>
        Underline = 4,

        /// <summary>
        ///     Strikethrough
        /// </summary>
        Strikethrough = 8,

        /// <summary>
        ///     Superscript
        /// </summary>
        Superscrit = 0x10,

        /// <summary>
        ///     Subscript
        /// </summary>
        Subscript = 0x20
    }

    internal interface IFont
    {
        string Name { get; set; }
        double Size { get; set; }
        FontStyles FontStyle { get; set; }
    }

    internal abstract class BaseFont : IFont
    {
        public string Name { get; set; }

        public double Size { get; set; }

        public FontStyles FontStyle { get; set; }
    }
}