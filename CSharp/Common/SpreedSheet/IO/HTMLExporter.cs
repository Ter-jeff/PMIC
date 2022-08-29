#define WPF

using System;
using System.IO;
using System.Text;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.XML;

namespace unvell.ReoGrid.IO
{
    internal class RGHTMLExporter
    {
        /// <summary>
        ///     Export grid as html5 into specified stream
        /// </summary>
        /// <param name="s">Stream contains the exported HTML5 content</param>
        /// <param name="sheet">Instance of worksheet</param>
        /// <param name="pageTitle">Custom page title of HTML page</param>
        /// <param name="htmlHeader">True to export default HTML header tag; false to export table content only</param>
        public static void Export(Stream s, Worksheet sheet, string pageTitle, bool htmlHeader = true)
        {
            using (var sw = new StreamWriter(s))
            {
                var sb = new StringBuilder();
                Cell cell;

                if (htmlHeader)
                {
                    sw.WriteLine("<!DOCTYPE html>");
                    sw.WriteLine("<html>");
                    sw.WriteLine("<head>");
                    sw.WriteLine("  <title>{0}</title>", pageTitle);
                    sw.WriteLine("  <meta content=\"text/html; charset=UTF-8\">");
                    sw.WriteLine("</head>");
                    sw.WriteLine("<body>");
                }

                sw.WriteLine("  <table style='border-collapse:collapse;border:none;'>");

                var maxRow = sheet.MaxContentRow;
                var maxCol = sheet.MaxContentCol;

                for (var r = 0; r <= maxRow; r++)
                {
                    var row = sheet.RetrieveRowHeader(r);

                    sw.WriteLine("    <tr style='height:{0}px;'>", row.InnerHeight);

                    for (var c = 0; c <= maxCol;)
                    {
                        var col = sheet.RetrieveColumnHeader(c);

                        cell = sheet.GetCell(r, c);

                        if (cell != null && (cell.Colspan <= 0 || cell.Rowspan <= 0))
                        {
                            c++;
                            continue;
                        }

                        sb.Length = 0;
                        sb.Append("      <td");

                        if (cell != null && cell.Rowspan > 1) sb.Append(" rowspan='" + cell.Rowspan + "'");
                        if (cell != null && cell.Colspan > 1) sb.Append(" colspan='" + cell.Colspan + "'");

                        sb.AppendFormat(" style='width:{0}px;", cell == null ? col.Width : cell.Width);

                        var halignOutputted = false;

                        if (cell != null)
                        {
                            // render horizontal align
                            if (cell.RenderHorAlign == GridRenderHorAlign.Right)
                            {
                                WriteHtmlStyle(sb, "text-align", "right");
                                halignOutputted = true;
                            }
                            else if (cell.RenderHorAlign == GridRenderHorAlign.Center)
                            {
                                WriteHtmlStyle(sb, "text-align", "center");
                                halignOutputted = true;
                            }
                        }

                        var style = sheet.GetCellStyles(r, c);
                        if (style != null)
                        {
                            // back color
                            if (style.HasStyle(PlainStyleFlag.BackColor) && style.BackColor != SolidColor.White)
                                WriteHtmlStyle(sb, "background-color", TextFormatHelper.EncodeColor(style.BackColor));

                            // text color
                            if (style.HasStyle(PlainStyleFlag.TextColor) && style.TextColor != SolidColor.Black)
                                WriteHtmlStyle(sb, "color", TextFormatHelper.EncodeColor(style.TextColor));

                            // font size
                            if (style.HasStyle(PlainStyleFlag.FontSize))
                                WriteHtmlStyle(sb, "font-size", style.FontSize + "pt");

                            // horizontal align
                            if (!halignOutputted && style.HasStyle(PlainStyleFlag.HorizontalAlign))
                                WriteHtmlStyle(sb, "text-align",
                                    XmlFileFormatHelper.EncodeHorizontalAlign(style.HAlign));

                            // vertical align
                            if (style.HasStyle(PlainStyleFlag.VerticalAlign))
                                WriteHtmlStyle(sb, "vertical-align",
                                    XmlFileFormatHelper.EncodeVerticalAlign(style.VAlign));
                        }

                        var rbi = sheet.GetRangeBorders(cell == null
                            ? new RangePosition(r, c, 1, 1)
                            : new RangePosition(cell.InternalRow, cell.InternalCol, cell.Rowspan, cell.Colspan));

                        if (!rbi.Top.IsEmpty) WriteCellBorder(sb, "border-top", rbi.Top);
                        if (!rbi.Left.IsEmpty) WriteCellBorder(sb, "border-left", rbi.Left);
                        if (!rbi.Right.IsEmpty) WriteCellBorder(sb, "border-right", rbi.Right);
                        if (!rbi.Bottom.IsEmpty) WriteCellBorder(sb, "border-bottom", rbi.Bottom);

                        sb.Append("'>");

                        sw.WriteLine(sb.ToString());

                        //cell = Grid.GetCell(r, c);

                        string text = null;
                        if (cell != null)
                            text = string.IsNullOrEmpty(cell.DisplayText)
                                    ? "&nbsp;"
                                    :
#if !CLIENT_PROFILE
                                    HtmlEncode(cell.DisplayText)
#else
								cell.DisplayText
#endif
                                // CLIENT_PROFILE
                                ;
                        else
                            text = "&nbsp;";

                        sw.WriteLine(text);

                        sw.WriteLine("      </td>");

                        c += cell == null ? 1 : cell.Colspan;
                    }

                    sw.WriteLine("    </tr>");
                }

                sw.WriteLine("  </table>");

                if (htmlHeader)
                {
                    sw.WriteLine("</body>");
                    sw.WriteLine("</html>");
                }
            }
        }

        /// <summary>
        ///     HTML-encodes a string and returns the encoded string.
        /// </summary>
        /// <remarks>
        ///     http://weblog.west-wind.com/posts/2009/Feb/05/Html-and-Uri-String-Encoding-without-SystemWeb
        /// </remarks>
        /// <param name="text">The text string to encode.</param>
        /// <returns>The HTML-encoded text.</returns>
        public static string HtmlEncode(string text)
        {
            if (text == null)
                return null;

            var sb = new StringBuilder(text.Length + (int)Math.Ceiling(text.Length * 0.3f));

            var len = text.Length;
            for (var i = 0; i < len; i++)
                switch (text[i])
                {
                    case '<':
                        sb.Append("&lt;");
                        break;
                    case '>':
                        sb.Append("&gt;");
                        break;
                    case '"':
                        sb.Append("&quot;");
                        break;
                    case '\'':
                        sb.Append("&apos;");
                        break;
                    case '&':
                        sb.Append("&amp;");
                        break;
                    default:
                        //if (text[i] > 159)
                        //{
                        //	// decimal numeric entity
                        //	sb.Append("&#");
                        //	sb.Append(((int)text[i]).ToString(System.Globalization.CultureInfo.InvariantCulture));
                        //	sb.Append(";");
                        //}
                        //else
                        sb.Append(text[i]);
                        break;
                }

            return sb.ToString();
        }

        private static void WriteHtmlStyle(StringBuilder sb, string name, string value)
        {
            sb.AppendFormat("{0}:{1};", name, value);
        }

        private static void WriteCellBorder(StringBuilder sb, string name, RangeBorderStyle borderStyle)
        {
            WriteHtmlStyle(sb, name, string.Format("{0} {1}",
                ToHTMLBorderLineStyle(borderStyle.Style), TextFormatHelper.EncodeColor(borderStyle.Color)));
        }

        private static string ToHTMLBorderLineStyle(BorderLineStyle borderLineStyle)
        {
            switch (borderLineStyle)
            {
                default:
                case BorderLineStyle.Solid:
                    return "solid 1px";
                case BorderLineStyle.Dashed:
                case BorderLineStyle.Dashed2:
                case BorderLineStyle.DashDotDot:
                case BorderLineStyle.DashDot:
                    return "dashed 1px";
                case BorderLineStyle.Dotted:
                    return "dotted 1px";
                case BorderLineStyle.BoldSolid:
                    return "solid 2px";
                case BorderLineStyle.BoldDashed:
                case BorderLineStyle.BoldDashDot:
                case BorderLineStyle.BoldDashDotDot:
                    return "dashed 2px";
                case BorderLineStyle.BoldDotted:
                    return "dotted 2px";
                case BorderLineStyle.BoldSolidStrong:
                    return "solid 3px";
            }
        }
    }
}