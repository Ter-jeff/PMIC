using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Referenced style for column header
    /// </summary>
    public class ColumnHeaderStyle : ReferenceStyle
    {
        /// <summary>
        ///     Column header instance
        /// </summary>
        private readonly ColumnHeader columnHeader;

        internal ColumnHeaderStyle(Worksheet grid, ColumnHeader columnHeader)
            : base(grid)
        {
            this.columnHeader = columnHeader;
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this row
        /// </summary>
        public bool Bold
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (bool)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.FontStyleBold);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
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
                CheckForReferenceOwner(columnHeader);

                return (bool)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.FontStyleItalic);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
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
                CheckForReferenceOwner(columnHeader);

                return (bool)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1,
                    PlainStyleFlag.FontStyleStrikethrough);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
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
                CheckForReferenceOwner(columnHeader);

                return (bool)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.FontStyleUnderline);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleUnderline,
                    Underline = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this column
        /// </summary>
        public GridHorAlign HorizontalAlign
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (GridHorAlign)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1,
                    PlainStyleFlag.HorizontalAlign);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.HorizontalAlign,
                    HAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set horizontal alignment for all cells on this column
        /// </summary>
        public GridVerAlign VerticalAlign
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (GridVerAlign)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1,
                    PlainStyleFlag.VerticalAlign);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.VerticalAlign,
                    VAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set padding for all cells on this column
        /// </summary>
        public PaddingValue Padding
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (PaddingValue)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.Padding);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.Padding,
                        Padding = value
                    });
            }
        }

        /// <summary>
        ///     Get or set background color for all cells on this column
        /// </summary>
        public SolidColor BackColor
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (SolidColor)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.BackColor);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor,
                        BackColor = value
                    });
            }
        }

        /// <summary>
        ///     Get or set background color for all cells on this column
        /// </summary>
        public SolidColor TextColor
        {
            get
            {
                CheckForReferenceOwner(columnHeader);

                return (SolidColor)Worksheet.GetRangeStyle(0, columnHeader.Index, -1, 1, PlainStyleFlag.TextColor);
            }
            set
            {
                CheckForReferenceOwner(columnHeader);

                //TODO: reduce create style object
                base.SetStyle(0, columnHeader.Index, -1, 1,
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.TextColor,
                        TextColor = value
                    });
            }
        }
    }
}