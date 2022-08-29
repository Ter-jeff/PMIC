using System;
using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Drawing.Text;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Styles of range or cells. By specifying PlainStyleFlag to determine
    ///     what styles should be used in this set.
    /// </summary>
    [Serializable]
    public class WorksheetRangeStyle
    {
        /// <summary>
        ///     Predefined empty style set.
        /// </summary>
        public static WorksheetRangeStyle Empty = new WorksheetRangeStyle();

        internal FontStyles fontStyles = FontStyles.Regular;

        /// <summary>
        ///     Create an empty style set.
        /// </summary>
        public WorksheetRangeStyle()
        {
        }

        /// <summary>
        ///     Create style set by copying from another one.
        /// </summary>
        /// <param name="source">Another style set to be copied.</param>
        public WorksheetRangeStyle(WorksheetRangeStyle source)
        {
            CopyFrom(source);
        }

        /// <summary>
        ///     Get or set the styles flag that indicates what styles are contained in this style set
        /// </summary>
        public PlainStyleFlag Flag { get; set; }

        /// <summary>
        ///     Get or set background color
        /// </summary>
        public SolidColor BackColor { get; set; }

        /// <summary>
        ///     Get or set backgrond pattern color.
        ///     When set pattern color or pattern style, the background color should also be set.
        /// </summary>
        public SolidColor FillPatternColor { get; set; }

        /// <summary>
        ///     Get or set background pattern style.
        ///     When set pattern color or pattern style, the background color should also be set.
        /// </summary>
        public HatchStyles FillPatternStyle { get; set; }

        /// <summary>
        ///     Get or set text color
        /// </summary>
        public SolidColor TextColor { get; set; }

        /// <summary>
        ///     Get or set font name
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        ///     Get or set font size
        /// </summary>
        public float FontSize { get; set; }

        /// <summary>
        ///     Get or set bold style
        /// </summary>
        public bool Bold
        {
            get { return (fontStyles & FontStyles.Bold) == FontStyles.Bold; }
            set
            {
                if (value)
                    fontStyles |= FontStyles.Bold;
                else
                    fontStyles &= ~FontStyles.Bold;
            }
        }

        /// <summary>
        ///     Get or set italic style
        /// </summary>
        public bool Italic
        {
            get { return (fontStyles & FontStyles.Italic) == FontStyles.Italic; }
            set
            {
                if (value)
                    fontStyles |= FontStyles.Italic;
                else
                    fontStyles &= ~FontStyles.Italic;
            }
        }

        /// <summary>
        ///     Get or set strikethrough style
        /// </summary>
        public bool Strikethrough
        {
            get { return (fontStyles & FontStyles.Strikethrough) == FontStyles.Strikethrough; }
            set
            {
                if (value)
                    fontStyles |= FontStyles.Strikethrough;
                else
                    fontStyles &= ~FontStyles.Strikethrough;
            }
        }

        /// <summary>
        ///     Get or set underline style
        /// </summary>
        public bool Underline
        {
            get { return (fontStyles & FontStyles.Underline) == FontStyles.Underline; }
            set
            {
                if (value)
                    fontStyles |= FontStyles.Underline;
                else
                    fontStyles &= ~FontStyles.Underline;
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment
        /// </summary>
        public GridHorAlign HAlign { get; set; }

        /// <summary>
        ///     Get or set vertical alignment
        /// </summary>
        public GridVerAlign VAlign { get; set; }

        /// <summary>
        ///     Get or set text-wrap mode
        /// </summary>
        public TextWrapMode TextWrapMode { get; set; }

        /// <summary>
        ///     Get or set text indent (0-65535)
        /// </summary>
        public ushort Indent { get; set; }

        /// <summary>
        ///     Get or set padding of cell.
        /// </summary>
        public PaddingValue Padding { get; set; }

        /// <summary>
        ///     Get or set rotate angle.
        /// </summary>
        public float RotationAngle { get; set; }

        /// <summary>
        ///     Clone style set from specified another style set.
        /// </summary>
        /// <param name="source">Another style to be cloned.</param>
        /// <returns>New cloned style set.</returns>
        public static WorksheetRangeStyle Clone(WorksheetRangeStyle source)
        {
            return source == null ? source : new WorksheetRangeStyle(source);
        }

        /// <summary>
        ///     Copy styles from another specified one.
        /// </summary>
        /// <param name="s">Style to be copied.</param>
        public void CopyFrom(WorksheetRangeStyle s)
        {
            StyleUtility.CopyStyle(s, this);
        }

        /// <summary>
        ///     Check two styles and compare whether or not they are same.
        /// </summary>
        /// <param name="obj">Another style object compared to this object.</param>
        /// <returns>True if thay are same; otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is WorksheetRangeStyle)) return false;

            var s2 = (WorksheetRangeStyle)obj;

            if (Flag != s2.Flag) return false;

            if ((Flag & PlainStyleFlag.HorizontalAlign) == PlainStyleFlag.HorizontalAlign
                && HAlign != s2.HAlign) return false;
            if ((Flag & PlainStyleFlag.VerticalAlign) == PlainStyleFlag.VerticalAlign
                && VAlign != s2.VAlign) return false;
            if ((Flag & PlainStyleFlag.BackColor) == PlainStyleFlag.BackColor
                && BackColor != s2.BackColor) return false;
            if ((Flag & PlainStyleFlag.FillPatternColor) == PlainStyleFlag.FillPatternColor
                && FillPatternColor != s2.FillPatternColor) return false;
            if ((Flag & PlainStyleFlag.FillPatternStyle) == PlainStyleFlag.FillPatternStyle
                && FillPatternStyle != s2.FillPatternStyle) return false;
            if ((Flag & PlainStyleFlag.TextColor) == PlainStyleFlag.TextColor
                && TextColor != s2.TextColor) return false;
            if ((Flag & PlainStyleFlag.FontName) == PlainStyleFlag.FontName
                && FontName != s2.FontName) return false;
            if ((Flag & PlainStyleFlag.FontSize) == PlainStyleFlag.FontSize
                && FontSize != s2.FontSize) return false;
            if ((Flag & PlainStyleFlag.FontStyleBold) == PlainStyleFlag.FontStyleBold
                && Bold != s2.Bold) return false;
            if ((Flag & PlainStyleFlag.FontStyleItalic) == PlainStyleFlag.FontStyleItalic
                && Italic != s2.Italic) return false;
            if ((Flag & PlainStyleFlag.FontStyleStrikethrough) == PlainStyleFlag.FontStyleStrikethrough
                && Strikethrough != s2.Strikethrough) return false;
            if ((Flag & PlainStyleFlag.FontStyleUnderline) == PlainStyleFlag.FontStyleUnderline
                && Underline != s2.Underline) return false;
            if ((Flag & PlainStyleFlag.TextWrap) == PlainStyleFlag.TextWrap
                && TextWrapMode != s2.TextWrapMode) return false;
            if ((Flag & PlainStyleFlag.Indent) == PlainStyleFlag.Indent
                && Indent != s2.Indent) return false;
            if ((Flag & PlainStyleFlag.Padding) == PlainStyleFlag.Padding
                && Padding != s2.Padding) return false;
            if ((Flag & PlainStyleFlag.RotationAngle) == PlainStyleFlag.RotationAngle
                && RotationAngle != s2.RotationAngle) return false;

            return true;
        }

        /// <summary>
        ///     Get hash code of this object.
        /// </summary>
        /// <returns>hash code</returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        ///     Check whether this set of style contains specified style item.
        /// </summary>
        /// <param name="flag">Style item to be checked.</param>
        /// <returns>Ture if this set contains specified style item.</returns>
        public bool HasStyle(PlainStyleFlag flag)
        {
            return (Flag & flag) == flag;
        }

        /// <summary>
        ///     Check whether this set of style contains any of one of specified style items.
        /// </summary>
        /// <param name="flag">Style items to be checked.</param>
        /// <returns>True if this set contains any one of specified items.</returns>
        public bool HasAny(PlainStyleFlag flag)
        {
            return (Flag & flag) > 0;
        }
    }
}