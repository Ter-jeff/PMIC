using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Refereced style for row header
    /// </summary>
    public class RowHeaderStyle : ReferenceStyle
    {
        private readonly RowHeader rowHeader;

        internal RowHeaderStyle(Worksheet grid, RowHeader rowHeader)
            : base(grid)
        {
            this.rowHeader = rowHeader;
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public bool Bold
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (bool)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.FontStyleBold);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleBold,
                    Bold = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public bool Italic
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (bool)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.FontStyleItalic);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleItalic,
                    Italic = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public bool Strikethrough
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (bool)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.FontStyleStrikethrough);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleStrikethrough,
                    Strikethrough = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public bool Underline
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (bool)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.FontStyleUnderline);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleUnderline,
                    Underline = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public GridHorAlign HorizontalAlign
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (GridHorAlign)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1,
                    PlainStyleFlag.HorizontalAlign);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.HorizontalAlign,
                    HAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public GridVerAlign VerticalAlign
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (GridVerAlign)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.VerticalAlign);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.VerticalAlign,
                    VAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set padding for all cells on this row
        /// </summary>
        public PaddingValue Padding
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (PaddingValue)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.Padding);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.Padding,
                        Padding = value
                    });
            }
        }

        /// <summary>
        ///     Get or set background color for all cells on this row
        /// </summary>
        public SolidColor BackColor
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (SolidColor)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.BackColor);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor,
                        BackColor = value
                    });
            }
        }

        /// <summary>
        ///     Get or set background color for all cells on this row
        /// </summary>
        public SolidColor TextColor
        {
            get
            {
                CheckForReferenceOwner(rowHeader);

                return (SolidColor)Worksheet.GetRangeStyle(rowHeader.Index, 0, 1, -1, PlainStyleFlag.TextColor);
            }
            set
            {
                CheckForReferenceOwner(rowHeader);

                //TODO: reduce create style object
                base.SetStyle(rowHeader.Index, 0, 1, -1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.TextColor,
                        TextColor = value
                    });
            }
        }
    }
}