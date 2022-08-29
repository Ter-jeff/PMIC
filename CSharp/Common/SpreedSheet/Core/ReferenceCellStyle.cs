using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    /// <summary>
    ///     Referenced cell style
    /// </summary>
    public class ReferenceCellStyle : ReferenceStyle
    {
        /// <summary>
        ///     Create referenced cell style
        /// </summary>
        /// <param name="cell"></param>
        public ReferenceCellStyle(Cell cell)
            : base(cell.Worksheet)
        {
            Cell = cell;
        }

        /// <summary>
        ///     Referenced cell instance
        /// </summary>
        public Cell Cell { get; }

        /// <summary>
        ///     Get or set cell background color.
        /// </summary>
        public SolidColor BackColor
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.BackColor;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.BackColor,
                    BackColor = value
                });
            }
        }

        /// <summary>
        ///     Get or set the horizontal alignment for the cell content.
        /// </summary>
        public GridHorAlign HAlign
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.HAlign;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.HorizontalAlign,
                    HAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set the vertical alignment for the cell content.
        /// </summary>
        public GridVerAlign VAlign
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.VAlign;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.VerticalAlign,
                    VAlign = value
                });
            }
        }

        /// <summary>
        ///     Get or set text color of cell.
        /// </summary>
        public SolidColor TextColor
        {
            get
            {
                CheckReferenceValidity();

                //SolidColor textColor;

                //if (!this.Cell.RenderColor.IsTransparent)
                //{
                //	// render color, used to render negative number, specified by data formatter
                //	textColor = this.Cell.RenderColor;
                //}
                //else if (this.Cell.InnerStyle.HasStyle(PlainStyleFlag.TextColor))
                //{
                //	// cell text color, specified by SetRangeStyle
                //	textColor = this.Cell.InnerStyle.TextColor;
                //}
                //// default cell text color
                //else if (!this.Cell.Worksheet.controlAdapter.ControlStyle.TryGetColor(
                //	ControlAppearanceColors.GridText, out textColor))
                //{
                //	// default built-in text
                //	textColor = SolidColor.Black;
                //}

                return Cell.InnerStyle.TextColor;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.TextColor,
                    TextColor = value
                });
            }
        }

        /// <summary>
        ///     Get or set font name of cell.
        /// </summary>
        public string FontName
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.FontName;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontName,
                    FontName = value
                });
            }
        }

        /// <summary>
        ///     Get or set font name of cell.
        /// </summary>
        public float FontSize
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.FontSize;
            }
            set
            {
                CheckReferenceValidity();

                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontSize,
                    FontSize = value
                });

                Worksheet.RequestInvalidate();
            }
        }

        /// <summary>
        ///     Determine whether or not the font style is bold.
        /// </summary>
        public bool Bold
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Bold;
            }
            set
            {
                CheckReferenceValidity();

                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleBold,
                    Bold = value
                });

                Worksheet.RequestInvalidate();
            }
        }

        /// <summary>
        ///     Determine whether or not the font style is italic.
        /// </summary>
        public bool Italic
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Italic;
            }
            set
            {
                CheckReferenceValidity();

                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleItalic,
                    Italic = value
                });

                Worksheet.RequestInvalidate();
            }
        }

        /// <summary>
        ///     Determine whether or not the font style has strikethrough.
        /// </summary>
        public bool Strikethrough
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Strikethrough;
            }
            set
            {
                CheckReferenceValidity();

                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleStrikethrough,
                    Strikethrough = value
                });

                Worksheet.RequestInvalidate();
            }
        }

        /// <summary>
        ///     Determine whether or not the font style has underline.
        /// </summary>
        public bool Underline
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Underline;
            }
            set
            {
                CheckReferenceValidity();

                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleUnderline,
                    Underline = value
                });
            }
        }

        /// <summary>
        ///     Get or set the cell text-wrap mode.
        /// </summary>
        public TextWrapMode TextWrap
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.TextWrapMode;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.TextWrap,
                    TextWrapMode = value
                });
            }
        }

        /// <summary>
        ///     Get or set cell indent.
        /// </summary>
        public ushort Indent
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Indent;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.Indent,
                    Indent = value
                });
            }
        }

        /// <summary>
        ///     Get or set padding of cell layout.
        /// </summary>
        public PaddingValue Padding
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.Padding;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.Padding,
                    Padding = value
                });
            }
        }

        /// <summary>
        ///     Get or set text rotation angle. (-90° ~ 90°)
        /// </summary>
        public float RotationAngle
        {
            get
            {
                CheckReferenceValidity();
                return Cell.InnerStyle.RotationAngle;
            }
            set
            {
                CheckReferenceValidity();
                Worksheet.SetCellStyleOwn(Cell, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.RotationAngle,
                    RotationAngle = value
                });
            }
        }

        private void CheckReferenceValidity()
        {
            if (Cell == null || Worksheet == null)
                throw new ReferenceObjectNotAssociatedException(
                    "ReferenceCellStyle must be associated to an valid cell and grid control.");
        }
    }
}