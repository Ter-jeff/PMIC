using System;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Referenced style instance to cell of range
    /// </summary>
    public abstract class ReferenceStyle
    {
        internal ReferenceStyle(Worksheet sheet)
        {
            Worksheet = sheet;
        }

        /// <summary>
        ///     Get worksheet instance
        /// </summary>
        public Worksheet Worksheet { get; }

        internal virtual void SetStyle(RangePosition range, WorksheetRangeStyle style)
        {
            SetStyle(range.Row, range.Col, range.Rows, range.Cols, style);
        }

        internal virtual void SetStyle(int row, int col, int rows, int cols, WorksheetRangeStyle style)
        {
            Worksheet.SetRangeStyles(row, col, rows, cols, style);
        }

        //TODO: reduce create style object
        internal virtual void SetStyle<T>(RangePosition range, PlainStyleFlag flag, T value)
        {
            SetStyle(range.Row, range.Col, range.Rows, range.Cols, flag, value);
        }

        //TODO: reduce create style object
        internal virtual void SetStyle<T>(int row, int col, int rows, int cols, PlainStyleFlag flag, T value)
        {
            var style = new WorksheetRangeStyle
            {
                Flag = flag
            };

            switch (flag)
            {
                case PlainStyleFlag.BackColor:
                    style.BackColor = (SolidColor)(object)value;
                    break;
                case PlainStyleFlag.TextColor:
                    style.TextColor = (SolidColor)(object)value;
                    break;
                case PlainStyleFlag.TextWrap:
                    style.TextWrapMode = (TextWrapMode)(object)value;
                    break;
                case PlainStyleFlag.Indent:
                    style.Indent = (ushort)(object)value;
                    break;
                case PlainStyleFlag.FillPatternColor:
                    style.FillPatternColor = (SolidColor)(object)value;
                    break;
                case PlainStyleFlag.FillPatternStyle:
                    style.FillPatternStyle = (HatchStyles)(object)value;
                    break;
                case PlainStyleFlag.Padding:
                    style.Padding = (PaddingValue)(object)value;
                    break;
                case PlainStyleFlag.FontName:
                    style.FontName = (string)(object)value;
                    break;
                case PlainStyleFlag.FontSize:
                    style.FontSize = (float)(object)value;
                    break;
                case PlainStyleFlag.FontStyleBold:
                    style.Bold = (bool)(object)value;
                    break;
                case PlainStyleFlag.FontStyleItalic:
                    style.Italic = (bool)(object)value;
                    break;
                case PlainStyleFlag.FontStyleUnderline:
                    style.Underline = (bool)(object)value;
                    break;
                case PlainStyleFlag.FontStyleStrikethrough:
                    style.Strikethrough = (bool)(object)value;
                    break;
                case PlainStyleFlag.HorizontalAlign:
                    style.HAlign = (GridHorAlign)(object)value;
                    break;
                case PlainStyleFlag.VerticalAlign:
                    style.VAlign = (GridVerAlign)(object)value;
                    break;
                case PlainStyleFlag.RotationAngle:
                    style.RotationAngle = (int)(object)value;
                    break;
            }

            Worksheet.SetRangeStyles(row, col, rows, cols, style);
        }

        internal virtual T GetStyle<T>(RangePosition range, PlainStyleFlag flag)
        {
            return GetStyle<T>(range.Row, range.Col, range.Rows, range.Cols, flag);
        }

        internal virtual T GetStyle<T>(int row, int col, int rows, int cols, PlainStyleFlag flag)
        {
            var type = typeof(T);

            return (T)Convert.ChangeType(Worksheet.GetRangeStyle(row, col, rows, cols, flag), type);
        }

        internal virtual void CheckForReferenceOwner(object owner)
        {
            if (Worksheet == null || owner == null)
                throw new ReferenceObjectNotAssociatedException(
                    "Reference style must be associated with an instance of owner.");
        }

        /// <summary>
        ///     Convert style reference to style object.
        /// </summary>
        /// <param name="refStyle">Style reference to be converted.</param>
        /// <returns>Style object converted from style reference.</returns>
        public static implicit operator WorksheetRangeStyle(ReferenceStyle refStyle)
        {
            return new WorksheetRangeStyle();
        }
    }
}