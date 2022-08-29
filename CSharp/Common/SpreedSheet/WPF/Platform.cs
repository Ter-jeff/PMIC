#define WPF

#if WPF

using System;
using System.Windows;
using System.Windows.Media;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.Common;
using unvell.Common.Win32Lib;
using unvell.ReoGrid.Graphics;
using FontStyles = unvell.ReoGrid.Drawing.Text.FontStyles;
using Size = unvell.ReoGrid.Graphics.Size;

namespace unvell.ReoGrid
{
    partial class Cell
    {
        [NonSerialized] internal FormattedText formattedText;
    }
}

namespace unvell.ReoGrid.Rendering
{
    #region PlatformUtility

    internal class PlatformUtility
    {
        private static double lastGetDPI;

        private static ResourcePoolManager resourcePoolManager; // = new ResourcePoolManager();

        internal static bool IsKeyDown(KeyCode key)
        {
            return Toolkit.IsKeyDown((Win32.VKey)key);
        }

        public static double GetDPI()
        {
            if (lastGetDPI == 0)
            {
                var win = new Window();
                var source = PresentationSource.FromVisual(win);

                if (source != null) lastGetDPI = 96.0 * source.CompositionTarget.TransformToDevice.M11;

                win.Close();

                if (lastGetDPI == 0) lastGetDPI = 96;
            }

            return lastGetDPI;
        }

        public static Typeface GetFontDefaultTypeface(FontFamily ff)
        {
            var typefaces = ff.GetTypefaces();
            if (typefaces.Count > 0)
            {
                var iterator = typefaces.GetEnumerator();
                if (iterator.MoveNext()) return iterator.Current;
            }

            return null;
        }

        internal static Size MeasureText(IRenderer r, string text, string fontName, double fontSize, FontStyles style)
        {
            ResourcePoolManager resManager;
            Typeface typeface = null;

            if (r == null)
            {
                if (resourcePoolManager == null) resourcePoolManager = new ResourcePoolManager();

                resManager = resourcePoolManager;
            }
            else
            {
                resManager = r.ResourcePoolManager;
            }

            typeface = resManager.GetTypeface(fontName, FontWeights.Regular, ToWPFFontStyle(style),
                FontStretches.Normal);

            if (typeface == null) return Size.Zero;

            //var typeface = new System.Windows.Media.Typeface(
            //			new System.Windows.Media.FontFamily(fontName),
            //			PlatformUtility.ToWPFFontStyle(this.fontStyles),
            //			(this.fontStyles & FontStyles.Bold) == FontStyles.Bold ?
            //			System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal,
            //			System.Windows.FontStretches.Normal);

            GlyphTypeface glyphTypeface;

            double totalWidth = 0;

            if (typeface.TryGetGlyphTypeface(out glyphTypeface))
            {
                //fontInfo.Ascent = typeface.FontFamily.Baseline;
                //fontInfo.LineHeight = typeface.CapsHeight;

                var size = fontSize * 1.33d;

                //this.GlyphIndexes.Capacity = text.Length;

                for (var n = 0; n < text.Length; n++)
                {
                    var glyphIndex = glyphTypeface.CharacterToGlyphMap[text[n]];
                    //GlyphIndexes.Add(glyphIndex);

                    var width = glyphTypeface.AdvanceWidths[glyphIndex] * size;
                    //this.TextSizes.Add(width);

                    totalWidth += width;
                }
            }

            return new Size(totalWidth, typeface.CapsHeight);
        }

        public static FontStyle ToWPFFontStyle(FontStyles textStyle)
        {
            if ((textStyle & FontStyles.Italic) == FontStyles.Italic)
                return System.Windows.FontStyles.Italic;
            return System.Windows.FontStyles.Normal;
        }

        public static TextDecorationCollection ToWPFFontDecorations(FontStyles textStyle)
        {
            var decorations = new TextDecorationCollection();

            if ((textStyle & FontStyles.Underline) == FontStyles.Underline) decorations.Add(TextDecorations.Underline);

            if ((textStyle & FontStyles.Strikethrough) == FontStyles.Strikethrough)
                decorations.Add(TextDecorations.Strikethrough);

            return decorations;
        }

        public static PenLineCap ToWPFLineCap(LineCapStyles capStyle)
        {
            switch (capStyle)
            {
                default:
                case LineCapStyles.None:
                    return PenLineCap.Flat;

                case LineCapStyles.Arrow:
                    return PenLineCap.Triangle;

                case LineCapStyles.Round:
                case LineCapStyles.Ellipse:
                    return PenLineCap.Round;
            }
        }
    }

    #endregion // PlatformUtility

    #region StaticResources

    internal class StaticResources
    {
        //private static string systemDefaultFontName = null;
        //internal static string SystemDefaultFontName
        //{
        //	get
        //	{
        //		if (systemDefaultFontName == null)
        //		{
        //			var names = System.Windows.SystemFonts.MessageFontFamily.FamilyNames;

        //			systemDefaultFontName = names.Count > 0 ? names[System.Windows.Markup.XmlLanguage.GetLanguage(string.Empty)] : string.Empty;
        //			//var typeface = ResourcePoolManager.Instance.GetTypeface(
        //		}

        //		return systemDefaultFontName;
        //	}
        //}
        //internal static double SystemDefaultFontSize = System.Drawing.SystemFonts.DefaultFont.Size * 72.0 / PlatformUtility.GetDPI();

        internal static readonly SolidColor SystemColor_Highlight = SystemColors.HighlightColor;
        internal static readonly SolidColor SystemColor_Window = SystemColors.WindowColor;
        internal static readonly SolidColor SystemColor_WindowText = SystemColors.WindowTextColor;
        internal static readonly SolidColor SystemColor_Control = SystemColors.ControlColor;
        internal static readonly SolidColor SystemColor_ControlLight = SystemColors.ControlLightColor;
        internal static readonly SolidColor SystemColor_ControlDark = SystemColors.ControlDarkColor;

        internal static readonly Pen Gray = new Pen(new SolidColorBrush(Colors.Gray), 1f);
    }

    #endregion // StaticResources
}

#endif