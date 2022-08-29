using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Range reference to spreadsheet
    /// </summary>
    public class ReferenceRangeStyle : ReferenceStyle
    {
        internal ReferenceRangeStyle(Worksheet grid, ReferenceRange range)
            : base(grid)
        {
            Range = range;
        }

        internal ReferenceRange Range { get; }

        /// <summary>
        ///     Get or set the background color to entire range
        /// </summary>
        public SolidColor BackColor
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<SolidColor>(Range, PlainStyleFlag.BackColor);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.BackColor, value);
            }
        }

        /// <summary>
        ///     Get or set the text color to entire range
        /// </summary>
        public SolidColor TextColor
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<SolidColor>(Range, PlainStyleFlag.TextColor);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.TextColor, value);
            }
        }

        /// <summary>
        ///     Get or set the font name to entire range
        /// </summary>
        public string FontName
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<string>(Range, PlainStyleFlag.FontName);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontName, value);
            }
        }

        /// <summary>
        ///     Get or set the font size to entire range
        /// </summary>
        public float FontSize
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<float>(Range, PlainStyleFlag.FontSize);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontSize, value);
            }
        }

        /// <summary>
        ///     Get or set bold font style to entire range
        /// </summary>
        public bool Bold
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<bool>(Range, PlainStyleFlag.FontStyleBold);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontStyleBold, value);
            }
        }

        /// <summary>
        ///     Get or set italic font style to entire range
        /// </summary>
        public bool Italic
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<bool>(Range, PlainStyleFlag.FontStyleItalic);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontStyleItalic, value);
            }
        }

        /// <summary>
        ///     Get or set underline font style to entire range
        /// </summary>
        public bool Underline
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<bool>(Range, PlainStyleFlag.FontStyleUnderline);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontStyleUnderline, value);
            }
        }

        /// <summary>
        ///     Get or set the strikethrough to entire range
        /// </summary>
        public bool Strikethrough
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<bool>(Range, PlainStyleFlag.FontStyleStrikethrough);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.FontStyleStrikethrough, value);
            }
        }

        /// <summary>
        ///     Get or set the horizontal alignment to entire range
        /// </summary>
        public GridHorAlign HorizontalAlign
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<GridHorAlign>(Range, PlainStyleFlag.HorizontalAlign);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.HorizontalAlign, value);
            }
        }

        /// <summary>
        ///     Get or set the vertical alignment to entire range
        /// </summary>
        public GridVerAlign VerticalAlign
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<GridVerAlign>(Range, PlainStyleFlag.VerticalAlign);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.VerticalAlign, value);
            }
        }

        /// <summary>
        ///     Get or set the padding to entire range
        /// </summary>
        public PaddingValue Padding
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<PaddingValue>(Range, PlainStyleFlag.Padding);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.Padding, value);
            }
        }

        /// <summary>
        ///     Get or set the text-wrap style to entire range
        /// </summary>
        public TextWrapMode TextWrap
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<TextWrapMode>(Range, PlainStyleFlag.TextWrap);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.TextWrap, value);
            }
        }

        /// <summary>
        ///     Get or set the cell indent
        /// </summary>
        public ushort Indent
        {
            get
            {
                CheckForReferenceOwner(Range);

                return GetStyle<ushort>(Range, PlainStyleFlag.Indent);
            }
            set
            {
                CheckForReferenceOwner(Range);

                SetStyle(Range, PlainStyleFlag.Padding, value);
            }
        }
    }
}