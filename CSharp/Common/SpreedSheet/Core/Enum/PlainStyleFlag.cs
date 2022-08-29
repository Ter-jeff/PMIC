namespace SpreedSheet.Core.Enum
{
    /// <summary>
    ///     Key of cell style item
    /// </summary>
    public enum PlainStyleFlag : long
    {
        /// <summary>
        ///     None style will be added or removed
        /// </summary>
        None = 0,

        /// <summary>
        ///     Font name
        /// </summary>
        FontName = 0x1,

        /// <summary>
        ///     Font size
        /// </summary>
        FontSize = 0x2,

        /// <summary>
        ///     Font bold
        /// </summary>
        FontStyleBold = 0x4,

        /// <summary>
        ///     Font italic
        /// </summary>
        FontStyleItalic = 0x8,

        /// <summary>
        ///     Font strikethrough
        /// </summary>
        FontStyleStrikethrough = 0x10,

        /// <summary>
        ///     Font underline
        /// </summary>
        FontStyleUnderline = 0x20,

        /// <summary>
        ///     Text color
        /// </summary>
        TextColor = 0x40,

        /// <summary>
        ///     Background color
        /// </summary>
        BackColor = 0x80,

        /// <summary>
        ///     Line color (Reserved)
        /// </summary>
        LineColor = 0x100,

        /// <summary>
        ///     Line style (Reserved)
        /// </summary>
        LineStyle = 0x200,

        /// <summary>
        ///     Line weight (Reserved)
        /// </summary>
        LineWeight = 0x400,

        /// <summary>
        ///     Line start cap (Reserved)
        /// </summary>
        LineStartCap = 0x800,

        /// <summary>
        ///     Line end cap (Reserved)
        /// </summary>
        LineEndCap = 0x1000,

        /// <summary>
        ///     Horizontal alignements
        /// </summary>
        HorizontalAlign = 0x2000,

        /// <summary>
        ///     Vertical alignement
        /// </summary>
        VerticalAlign = 0x4000,

        /// <summary>
        ///     Background pattern color (not supported in WPF version)
        /// </summary>
        FillPatternColor = 0x80000,

        /// <summary>
        ///     Background pattern style (not supported in WPF version)
        /// </summary>
        FillPatternStyle = 0x100000,

        /// <summary>
        ///     Text wrap (word-break mode)
        /// </summary>
        TextWrap = 0x200000,

        /// <summary>
        ///     Padding
        /// </summary>
        Indent = 0x400000,

        /// <summary>
        ///     Padding
        /// </summary>
        Padding = 0x800000,

        /// <summary>
        ///     Rotation angle for cell text
        /// </summary>
        RotationAngle = 0x1000000,

        /// <summary>
        ///     [Union flag] All flags of font style
        /// </summary>
        FontStyleAll = FontStyleBold | FontStyleItalic
                                     | FontStyleStrikethrough | FontStyleUnderline,

        /// <summary>
        ///     [Union flag] All font styles (name, size and style)
        /// </summary>
        FontAll = FontName | FontSize | FontStyleAll,

        /// <summary>
        ///     [Union flag] All line styles (color, style, weight and caps)
        /// </summary>
        LineAll = LineColor | LineStyle | LineWeight | LineStartCap | LineEndCap,

        /// <summary>
        ///     [Union flag] All layout styles (Text-wrap, padding and angle)
        /// </summary>
        LayoutAll = TextWrap | Padding | RotationAngle,

        /// <summary>
        ///     [Union flag] Both horizontal and vertical alignments
        /// </summary>
        AlignAll = HorizontalAlign | VerticalAlign,

        /// <summary>
        ///     [Union flag] Background pattern (color and style)
        /// </summary>
        FillPattern = FillPatternColor | FillPatternStyle,

        /// <summary>
        ///     [Union flag] All background styles (color and pattern)
        /// </summary>
        BackAll = BackColor | FillPattern,

        /// <summary>
        ///     [Union flag] All styles
        /// </summary>
        All = FontAll | TextColor | BackAll | LineAll | AlignAll | LayoutAll
    }
}