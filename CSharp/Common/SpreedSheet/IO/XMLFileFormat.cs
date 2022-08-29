#define WPF

// Disable XML comment document
#pragma warning disable 1591

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid.Core;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Graphics;
#if PRINT
using unvell.ReoGrid.Print;
#endif // PRINT

namespace unvell.ReoGrid
{
    // Classes in this namespace used to persistence the grid control.
    // Data of cells, Styles and Borders of range, Script and Macro,
    // all the data belonging to an instance of grid control will be 
    // converted into XML and stored in specified stream.
    namespace XML
    {
        internal sealed class XmlFileFormatHelper
        {
            #region Alignement

            internal static string EncodeHorizontalAlign(GridHorAlign halign)
            {
                switch (halign)
                {
                    default:
                    case GridHorAlign.General:
                        return "general";
                    case GridHorAlign.Left:
                        return "left";
                    case GridHorAlign.Center:
                        return "center";
                    case GridHorAlign.Right:
                        return "right";
                    case GridHorAlign.DistributedIndent:
                        return "distributed-indent";
                }
            }

            internal static string EncodeVerticalAlign(GridVerAlign valign)
            {
                switch (valign)
                {
                    case GridVerAlign.Top:
                        return "top";
                    default:
                    case GridVerAlign.Middle:
                        return "middle";
                    case GridVerAlign.Bottom:
                        return "bottom";
                }
            }

            internal static GridHorAlign DecodeHorizontalAlign(string align)
            {
                switch (align)
                {
                    default:
                    case "general":
                        return GridHorAlign.General;
                    case "left":
                        return GridHorAlign.Left;
                    case "center":
                        return GridHorAlign.Center;
                    case "right":
                        return GridHorAlign.Right;
                    case "distributed-indent":
                        return GridHorAlign.DistributedIndent;
                }
            }

            internal static GridVerAlign DecodeVerticalAlign(string valign)
            {
                switch (valign)
                {
                    case "top":
                        return GridVerAlign.Top;
                    default:
                    case "middle":
                        return GridVerAlign.Middle;
                    case "bottom":
                        return GridVerAlign.Bottom;
                }
            }

            #endregion

            #region Data Format

            internal static string EncodeCellDataFormat(CellDataFormatFlag format)
            {
                switch (format)
                {
                    default:
                    case CellDataFormatFlag.General:
                        return null;
                    case CellDataFormatFlag.Number:
                        return "number";
                    case CellDataFormatFlag.Text:
                        return "text";
                    case CellDataFormatFlag.DateTime:
                        return "datetime";
                    case CellDataFormatFlag.Percent:
                        return "percent";
                    case CellDataFormatFlag.Currency:
                        return "currency";
                }
            }

            internal static CellDataFormatFlag DecodeCellDataFormat(string format)
            {
                if (format == null) return CellDataFormatFlag.General;

                switch (format.ToLower())
                {
                    default:
                        return CellDataFormatFlag.General;
                    case "number":
                        return CellDataFormatFlag.Number;
                    case "text":
                        return CellDataFormatFlag.Text;
                    case "datetime":
                        return CellDataFormatFlag.DateTime;
                    case "percent":
                        return CellDataFormatFlag.Percent;
                    case "currency":
                        return CellDataFormatFlag.Currency;
                }
            }

            internal static string EncodeNegativeNumberStyle(
                NumberDataFormatter.NumberNegativeStyle numberNegativeStyle)
            {
                if (numberNegativeStyle == NumberDataFormatter.NumberNegativeStyle.Default) return null;

                var sb = new StringBuilder(30);

                if ((numberNegativeStyle & NumberDataFormatter.NumberNegativeStyle.Red) ==
                    NumberDataFormatter.NumberNegativeStyle.Red)
                    sb.Append("red ");
                if ((numberNegativeStyle & NumberDataFormatter.NumberNegativeStyle.Brackets) ==
                    NumberDataFormatter.NumberNegativeStyle.Brackets)
                    sb.Append("brackets ");
                if ((numberNegativeStyle & NumberDataFormatter.NumberNegativeStyle.Prefix_Sankaku) ==
                    NumberDataFormatter.NumberNegativeStyle.Prefix_Sankaku)
                    sb.Append("sankaku ");
                if ((numberNegativeStyle & NumberDataFormatter.NumberNegativeStyle.CustomSymbol) ==
                    NumberDataFormatter.NumberNegativeStyle.CustomSymbol)
                    sb.Append("custom ");

                if (sb[sb.Length - 1] == ' ') sb.Length--;

                return sb.ToString();
            }

            internal static NumberDataFormatter.NumberNegativeStyle DecodeNegativeNumberStyle(string p)
            {
                var flag = NumberDataFormatter.NumberNegativeStyle.Default;

                if (string.IsNullOrEmpty(p)) return flag;

                var hasMinusStyle = false;

                var tokens = p.Split(' ');

                foreach (var token in tokens)
                    switch (token)
                    {
                        case "red":
                            flag |= NumberDataFormatter.NumberNegativeStyle.Red;
                            break;
                        case "brackets":
                            flag |= NumberDataFormatter.NumberNegativeStyle.Brackets;
                            break;
                        case "minus":
                            hasMinusStyle = true;
                            break;
                        case "sankaku":
                            flag |= NumberDataFormatter.NumberNegativeStyle.Prefix_Sankaku;
                            break;
                        case "custom":
                            flag |= NumberDataFormatter.NumberNegativeStyle.CustomSymbol;
                            break;
                    }

                if (!hasMinusStyle) flag &= ~NumberDataFormatter.NumberNegativeStyle.Minus;

                return flag;
            }

            #endregion

            #region Border

            internal static string EncodeBorderPos(BorderPositions pos)
            {
                return pos.ToString().ToLower();
            }

            internal static object DecodeBorderPos(string p)
            {
                if (string.IsNullOrEmpty(p)) return BorderPositions.None;
                // TODO: need convert first char upper 
                return (BorderPositions)Enum.Parse(typeof(BorderPositions), p);
            }

            internal static string EncodeHBorderOwnerPos(HBorderOwnerPosition pos)
            {
                return pos.ToString().ToLower();
            }

            internal static HBorderOwnerPosition DecodeHBorderOwnerPos(string p)
            {
                if (string.IsNullOrEmpty(p)) return HBorderOwnerPosition.None;
                if (p.Equals("all", StringComparison.CurrentCultureIgnoreCase))
                    return HBorderOwnerPosition.All;
                if (p.Equals("top", StringComparison.CurrentCultureIgnoreCase))
                    return HBorderOwnerPosition.Top;
                if (p.Equals("bottom", StringComparison.CurrentCultureIgnoreCase))
                    return HBorderOwnerPosition.Bottom;
                return HBorderOwnerPosition.None;
            }

            internal static string EncodeVBorderOwnerPos(VBorderOwnerPosition pos)
            {
                return pos.ToString().ToLower();
            }

            internal static VBorderOwnerPosition DecodeVBorderOwnerPos(string p)
            {
                if (string.IsNullOrEmpty(p)) return VBorderOwnerPosition.None;
                if (p.Equals("all", StringComparison.CurrentCultureIgnoreCase))
                    return VBorderOwnerPosition.All;
                if (p.Equals("left", StringComparison.CurrentCultureIgnoreCase))
                    return VBorderOwnerPosition.Left;
                if (p.Equals("right", StringComparison.CurrentCultureIgnoreCase))
                    return VBorderOwnerPosition.Right;
                return VBorderOwnerPosition.None;
            }

            #endregion

            #region TextWrap

            internal static string EncodeTextWrapMode(TextWrapMode wrapMode)
            {
                switch (wrapMode)
                {
                    default:
                    case TextWrapMode.NoWrap:
                        return "no-wrap";
                    case TextWrapMode.WordBreak:
                        return "word-break";
                    case TextWrapMode.BreakAll:
                        return "break-all";
                }
            }

            internal static TextWrapMode DecodeTextWrapMode(string p)
            {
                if (string.IsNullOrEmpty(p)) return TextWrapMode.NoWrap;

                if (p.Equals("word-break", StringComparison.CurrentCultureIgnoreCase))
                    return TextWrapMode.WordBreak;
                if (p.Equals("break-all", StringComparison.CurrentCultureIgnoreCase))
                    return TextWrapMode.BreakAll;
                return TextWrapMode.NoWrap;
            }

            #endregion

            #region Selection Mode & Style & Forward Direction

            internal static string EncodeSelectionMode(WorksheetSelectionMode selMode)
            {
                switch (selMode)
                {
                    default:
                    case WorksheetSelectionMode.Range: return "range";
                    case WorksheetSelectionMode.Cell: return "cell";
                    case WorksheetSelectionMode.None: return "none";
                    case WorksheetSelectionMode.Row: return "row";
                    case WorksheetSelectionMode.Column: return "column";
                }
            }

            internal static WorksheetSelectionMode DecodeSelectionMode(string arg)
            {
                if (arg.Equals("cell", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionMode.Cell;
                if (arg.Equals("none", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionMode.None;
                if (arg.Equals("row", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionMode.Row;
                if (arg.Equals("column", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionMode.Column;
                return WorksheetSelectionMode.Range;
            }

            internal static string EncodeSelectionStyle(WorksheetSelectionStyle selStyle)
            {
                switch (selStyle)
                {
                    default:
                    case WorksheetSelectionStyle.Default: return "default";
                    case WorksheetSelectionStyle.FocusRect: return "windows-focus";
                    case WorksheetSelectionStyle.None: return "none";
                }
            }

            internal static WorksheetSelectionStyle DecodeSelectionStyle(string arg)
            {
                if (arg.Equals("windows-focus", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionStyle.FocusRect;
                if (arg.Equals("default", StringComparison.CurrentCultureIgnoreCase))
                    return WorksheetSelectionStyle.None;
                return WorksheetSelectionStyle.Default;
            }

            internal static string EncodeFocusForwardDirection(SelectionForwardDirection forwardDirection)
            {
                switch (forwardDirection)
                {
                    default:
                    case SelectionForwardDirection.Right: return "right";
                    case SelectionForwardDirection.Down: return "down";
                }
            }

            internal static SelectionForwardDirection DecodeFocusForwardDirection(string arg)
            {
                if (arg.Equals("down", StringComparison.CurrentCultureIgnoreCase))
                    return SelectionForwardDirection.Down;
                return SelectionForwardDirection.Right;
            }

            internal static string EncodeFocusPosStyle(FocusPosStyle focusPosStyle)
            {
                switch (focusPosStyle)
                {
                    default:
                    case FocusPosStyle.Default: return "default";
                    case FocusPosStyle.None: return "none";
                }
            }

            internal static FocusPosStyle DecodeFocusPosStyle(string arg)
            {
                if (arg.Equals("none", StringComparison.CurrentCultureIgnoreCase))
                    return FocusPosStyle.None;
                return FocusPosStyle.Default;
            }

            #endregion // Selection Mode & Style & Forward Direction

            #region Print

#if PRINT
			internal static string EncodePageOrder(PrintPageOrder pageOrder)
			{
				switch (pageOrder)
				{
					default:
					case PrintPageOrder.DownThenOver:
						return "down-over";
					case PrintPageOrder.OverThenDown:
						return "over-down";
				}
			}

			internal static PrintPageOrder DecodePageOrder(string data)
			{
				if (data.Equals("over-down", StringComparison.CurrentCultureIgnoreCase))
					return PrintPageOrder.OverThenDown;
				else // default
					return PrintPageOrder.DownThenOver;
			}
#endif // PRINT

            #endregion // Print

            public static string EncodeFreezeArea(FreezeArea area)
            {
                switch (area)
                {
                    default:
                    case FreezeArea.LeftTop: return "left-top";
                    case FreezeArea.LeftBottom: return "left-top";
                    case FreezeArea.RightTop: return "left-top";
                    case FreezeArea.RightBottom: return "left-top";
                }
            }

            public static FreezeArea DecodeFreezeArea(string str)
            {
                if (string.IsNullOrEmpty(str))
                    return FreezeArea.LeftTop;
                if (string.Compare(str, "left-top", true) == 0)
                    return FreezeArea.LeftTop;
                if (string.Compare(str, "left-bottom", true) == 0)
                    return FreezeArea.LeftBottom;
                if (string.Compare(str, "right-top", true) == 0)
                    return FreezeArea.RightTop;
                if (string.Compare(str, "right-bottom", true) == 0)
                    return FreezeArea.RightBottom;
                return FreezeArea.LeftTop;
            }
        }

        #region XML Serialization

        [Obfuscation(Feature = "renaming", Exclude = true)]
        [XmlRoot("grid")]
        public class RGXmlSheet
        {
            [XmlArray("cells")] [XmlArrayItem("cell")]
            public List<RGXmlCell> cells = new List<RGXmlCell>();

            [XmlArray("cols")] [XmlArrayItem("col")]
            public List<RGXmlColHead> cols = new List<RGXmlColHead>();

            [XmlArray("h-borders")] [XmlArrayItem("h-border")]
            public List<RGXmlHBorder> hborder = new List<RGXmlHBorder>();

            [XmlElement("head")] public RGXmlHead head;

            [XmlArray("rows")] [XmlArrayItem("row")]
            public List<RGXmlRowHead> rows = new List<RGXmlRowHead>();

            [XmlElement("style")] public RGXmlCellStyle style;

            [XmlArray("v-borders")] [XmlArrayItem("v-border")]
            public List<RGXmlVBorder> vborder = new List<RGXmlVBorder>();
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlHead
        {
            [XmlElement("meta")] public RGXmlMeta meta;
            [XmlElement("rows")] public int rows;
            [XmlElement("cols")] public int cols;

            [XmlElement("default-row-height")] public ushort defaultRowHeight;
            [XmlElement("default-col-width")] public ushort defaultColumnWidth;


            [XmlElement("selection-mode")] public string selectionMode;
            [XmlElement("selection-style")] public string selectionStyle;

            [XmlElement("focus-forward-direction")]
            public string focusForwardDirection;

            [XmlElement("focus-cell-style")] public string focusCellStyle;


            [XmlElement("row-header-panel-width")] public string rowHeaderWidth;
            [XmlElement("freeze-row")] public string freezeRow;
            [XmlElement("freeze-col")] public string freezeCol;
            [XmlElement("freeze-area")] public string freezeArea;


            [XmlArray("named-ranges")] [XmlArrayItem("named-range")]
            public List<RGXmlNamedRange> namedRanges;

            [XmlElement("outlines")] public RGXmlOutlineList outlines;

            [XmlElement("print-settings")] public RGXmlPrintSetting printSettings;


            [XmlElement("settings")] public RGXmlWorksheetSetting settings;
            [XmlElement("script")] public RGXmlScript script;


            //[Obsolete("moved into meta element")]
            //[XmlElement("culture")]
            //public string culture;

            //[Obsolete("moved into meta element")]
            //[XmlElement("editor")]
            //public string editor;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlMeta
        {
            [XmlElement("culture")] public string culture;
            [XmlElement("editor")] public string editor;
            [XmlElement("core-ver")] public string controlVersion;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlPrintSetting
        {
            [XmlElement("page-break-cols")] public string pageBreakCols;
            [XmlElement("scale-factor")] public string scaling;
            [XmlElement("paper-size")] public string paperName;
            [XmlElement("landscape")] public string landscape;

            [XmlElement("page-order")] public string pageOrder;
            [XmlElement("paper-width")] public string paperWidth;
            [XmlElement("paper-height")] public string paperHeight;
            [XmlElement("margins")] public RGXmlMargins margins;


            [XmlElement("page-break-rows")] public string pageBreakRows;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlPageBreaks
        {
            [XmlAttribute("col")] public int col;

            [XmlAttribute("row")] public int row;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlMargins
        {
            [XmlElement("right")] public double right;
            [XmlElement("top")] public double top;
            [XmlElement("bottom")] public double bottom;

            [XmlElement("left")] public double left;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlWorksheetSetting
        {
            [XmlElement("selection-mode")] public string selectionMode;
            [XmlElement("selection-style")] public string selectionStyle;
            [XmlElement("readonly")] public string @readonly;

            [XmlElement("allow-adjust-row-height")]
            public string allowAdjustRowHeight;

            [XmlElement("allow-adjust-column-width")]
            public string allowAdjustColumnWidth;


            [XmlAttribute("meta")] public string metaValue;


            [XmlElement("show-page-break")] public string showPageBreakes;
            [XmlElement("show-row-header")] public string showRowHeader;
            [XmlElement("show-col-header")] public string showColHeader;

            [XmlElement("show-grid")] public string showGrid;


            [XmlElement("show-h-scrollbar")] public string showHScrollBar;

            [XmlElement("show-v-scrollbar")] public string showVScrollBar;
        }

        [Obfuscation(Feature = "renaming", Exclude = true)]
        public class RGXmlRowHead
        {
            [XmlAttribute("auto-height")] public string autoHeight;

            [XmlAttribute] public ushort height;

            [XmlAttribute("last-height")] public string lastHeight;

            [XmlAttribute] public int row;

            [XmlElement("style")] public RGXmlCellStyle style;

            [XmlAttribute] public string text;

            [XmlAttribute("text-color")] public string textColor;
        }

        public class RGXmlColHead
        {
            [XmlAttribute("auto-width")] public string autoWidth;

            [XmlAttribute] public int col;

            [XmlAttribute("cell-body-type")] public string defaultCellBody;

            [XmlAttribute("last-width")] public string lastWidth;

            [XmlElement("style")] public RGXmlCellStyle style;

            [XmlAttribute] public string text;

            [XmlAttribute("text-color")] public string textColor;

            [XmlAttribute] public ushort width;
        }

        public class RGXmlVBorder : RGXmlBorder
        {
            [XmlAttribute] public int rows;

            public RGXmlVBorder()
            {
            }

            internal RGXmlVBorder(int row, int col, int rows, RangeBorderStyle borderStyle, VBorderOwnerPosition pos)
                : base(row, col, borderStyle, XmlFileFormatHelper.EncodeVBorderOwnerPos(pos))
            {
                this.rows = rows;
            }
        }

        public class RGXmlHBorder : RGXmlBorder
        {
            [XmlAttribute] public int cols;

            public RGXmlHBorder()
            {
            }

            internal RGXmlHBorder(int row, int col, int cols, RangeBorderStyle borderStyle, HBorderOwnerPosition pos)
                : base(row, col, borderStyle, XmlFileFormatHelper.EncodeHBorderOwnerPos(pos))
            {
                this.cols = cols;
            }
        }

        public class RGXmlBorder
        {
            [XmlAttribute] public int col;

            [XmlAttribute] public string color;

            [XmlAttribute] public string pos;

            [XmlAttribute] public int row;

            [XmlAttribute] public string style;

            public RGXmlBorder()
            {
            }

            internal RGXmlBorder(int row, int col, RangeBorderStyle borderStyle, string pos)
            {
                this.row = row;
                this.col = col;
                this.pos = pos;

                if (borderStyle.Color != SolidColor.Black) color = TextFormatHelper.EncodeColor(borderStyle.Color);
                if (borderStyle.Style != BorderLineStyle.Solid) style = borderStyle.Style.ToString();
            }

            [XmlIgnore]
            internal RangeBorderStyle StyleGridBorder
            {
                get
                {
                    BorderLineStyle borderStyle;

                    if (string.IsNullOrEmpty(style))
                    {
                        borderStyle = BorderLineStyle.Solid;
                    }
                    else
                    {
                        if (style.Equals("dot")) // for Backward-compatibility
                            borderStyle = BorderLineStyle.Dotted;
                        else if (style.Equals("dash")) // for Backward-compatibility
                            borderStyle = BorderLineStyle.Dashed;
                        else
                            borderStyle = (BorderLineStyle)Enum.Parse(typeof(BorderLineStyle), style, true);
                    }

                    SolidColor rgbColor;

                    if (string.IsNullOrEmpty(color))
                        rgbColor = SolidColor.Black;
                    else if (!TextFormatHelper.DecodeColor(color, out rgbColor)) rgbColor = SolidColor.Black;

                    return new RangeBorderStyle
                    {
                        Color = rgbColor,
                        Style = borderStyle
                    };
                }
            }
        }

        public class RGXmlCell
        {
            [XmlAttribute("body-type")] public string bodyType;

            [XmlAttribute] public int col;

            [XmlAttribute] public string colspan;

            [XmlText] public string data;

            [XmlAttribute("format")] public string dataFormat;

            [XmlElement] public RGXmlCellFormual formula;
            [XmlElement("format-args")] public RGXmlCellDataFormatArgs dataFormatArgs;


            // add by Rick
            [XmlAttribute("readonly")] public string @readonly;

            [XmlAttribute] public int row;

            [XmlAttribute] public string rowspan;

            [XmlElement("style")] public RGXmlCellStyle style;

            [XmlAttribute("trace-dependents")] public string traceDependents;

            [XmlAttribute("trace-precedents")] public string tracePrecedents;
        }

        public class RGXmlCellFormual
        {
            [XmlText] public string val;
        }

        public class RGXmlCellStyle
        {
            [XmlAttribute("bgcolor")] public string backColor;

            [XmlAttribute("bold")] public string bold;

            [XmlElement("fill-pattern")] public RGXmlCellStyleFillPattern fillPattern;

            [XmlAttribute("font")] public string font;

            [XmlAttribute("font-size")] public string fontSize;

            [XmlAttribute("align")] public string hAlign;

            [XmlAttribute("indent")] public string indent;

            [XmlAttribute("italic")] public string italic;

            [XmlAttribute("padding")] public string padding;

            [XmlAttribute("rotate-angle")] public string rotateAngle;

            [XmlAttribute("strikethrough")] public string strikethrough;

            [XmlAttribute("color")] public string textColor;

            [XmlAttribute("text-wrap")] public string textWrap;

            [XmlAttribute("underline")] public string underline;

            [XmlAttribute("valign")] public string vAlign;
        }

        public class RGXmlCellStyleFillPattern
        {
            [XmlAttribute("color")] public string color;

            [XmlAttribute("style")] public int patternStyleId;
        }

        public class RGXmlCellDataFormatArgs
        {
            [XmlAttribute("culture")] public string culture;

            [XmlAttribute("decimal-places")] public string decimalPlaces;

            [XmlAttribute("negative-style")] public string negativeStyle;

            [XmlAttribute("pattern")] public string pattern;

            [XmlAttribute("use-separator")] public string useSeparator;

            public bool IsEmpty
            {
                get
                {
                    return string.IsNullOrEmpty(decimalPlaces)
                           && string.IsNullOrEmpty(useSeparator)
                           && string.IsNullOrEmpty(pattern)
                           && string.IsNullOrEmpty(culture)
                           && string.IsNullOrEmpty(negativeStyle);
                }
            }
        }

        public class RGXmlOutlineList
        {
            [XmlElement("row-outline")] public List<RGXmlOutline> rowOutlines;
            [XmlElement("col-outline")] public List<RGXmlOutline> colOutlines;
        }

        public class RGXmlOutline
        {
            [XmlAttribute("collapsed")] public bool collapsed;

            [XmlAttribute("count")] public int count;

            [XmlAttribute("start")] public int start;
        }

        public class RGXmlNamedRange
        {
            [XmlAttribute("address")] public string address;

            [XmlText] public string comment;

            [XmlAttribute("name")] public string name;
        }

        public class RGXmlScript
        {
            [XmlText] public string content;
        }

        #endregion // XML
    }
}