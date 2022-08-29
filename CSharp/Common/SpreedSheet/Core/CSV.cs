#define WPF

using System;
using System.IO;
using System.Text;
using SpreedSheet.Core;
using SpreedSheet.Interaction;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.IO;

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        #region Load

        /// <summary>
        ///     Load CSV file into worksheet.
        /// </summary>
        /// <param name="path">File contains CSV data.</param>
        public void LoadCSV(string path)
        {
            LoadCSV(path, RangePosition.EntireRange);
        }

        /// <summary>
        ///     Load CSV file into worksheet.
        /// </summary>
        /// <param name="path">File contains CSV data.</param>
        /// <param name="targetRange">The range used to fill loaded CSV data.</param>
        public void LoadCSV(string path, RangePosition targetRange)
        {
            LoadCSV(path, Encoding.Default, targetRange);
        }

        /// <summary>
        ///     Load CSV file into worksheet.
        /// </summary>
        /// <param name="path">Path to load CSV file.</param>
        /// <param name="encoding">Encoding used to read and decode plain-text from file.</param>
        public void LoadCSV(string path, Encoding encoding)
        {
            LoadCSV(path, encoding, RangePosition.EntireRange);
        }

        /// <summary>
        ///     Load CSV file into worksheet.
        /// </summary>
        /// <param name="path">Path to load CSV file.</param>
        /// <param name="encoding">Encoding used to read and decode plain-text from file.</param>
        /// <param name="targetRange">The range used to fill loaded CSV data.</param>
        public void LoadCSV(string path, Encoding encoding, RangePosition targetRange)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                LoadCSV(fs, encoding, targetRange);
            }
        }

        /// <summary>
        ///     Load CSV data from stream into worksheet.
        /// </summary>
        /// <param name="s">Input stream to read CSV data.</param>
        public void LoadCSV(Stream s)
        {
            LoadCSV(s, Encoding.Default);
        }

        /// <summary>
        ///     Load CSV data from stream into worksheet.
        /// </summary>
        /// <param name="s">Input stream to read CSV data.</param>
        /// <param name="targetRange">The range used to fill loaded CSV data.</param>
        public void LoadCSV(Stream s, RangePosition targetRange)
        {
            LoadCSV(s, Encoding.Default, targetRange);
        }

        /// <summary>
        ///     Load CSV data from stream into worksheet.
        /// </summary>
        /// <param name="s">Input stream to read CSV data.</param>
        /// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
        public void LoadCSV(Stream s, Encoding encoding)
        {
            LoadCSV(s, encoding, RangePosition.EntireRange);
        }

        /// <summary>
        ///     Load CSV data from stream into worksheet.
        /// </summary>
        /// <param name="s">Input stream to read CSV data.</param>
        /// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
        /// <param name="targetRange">The range used to fill loaded CSV data.</param>
        public void LoadCSV(Stream s, Encoding encoding, RangePosition targetRange)
        {
            LoadCSV(s, encoding, targetRange, targetRange.IsEntire, 256);
        }

        /// <summary>
        ///     Load CSV data from stream into worksheet.
        /// </summary>
        /// <param name="s">Input stream to read CSV data.</param>
        /// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
        /// <param name="targetRange">The range used to fill loaded CSV data.</param>
        /// <param name="autoSpread">decide whether or not to append rows or columns automatically to fill csv data</param>
        /// <param name="bufferLines">decide how many lines int the buffer to read and fill csv data</param>
        public void LoadCSV(Stream s, Encoding encoding, RangePosition targetRange, bool autoSpread, int bufferLines)
        {
            controlAdapter?.ChangeCursor(CursorStyle.Busy);

            try
            {
                var csvProvider = new CSVFileFormatProvider();

                var arg = new CSVFormatArgument
                {
                    AutoSpread = autoSpread,
                    BufferLines = bufferLines,
                    TargetRange = targetRange
                };

                Clear();

                csvProvider.Load(workbook, s, encoding, arg, "");
            }
            finally
            {
                controlAdapter?.ChangeCursor(CursorStyle.PlatformDefault);
            }
        }

        #endregion // Load

        #region Export

        /// <summary>
        ///     Export spreadsheet as CSV format from specified number of rows.
        /// </summary>
        /// <param name="path">File path to write CSV format as stream.</param>
        /// <param name="startRow">
        ///     Number of rows start to export data,
        ///     this property is useful to skip the headers on top of worksheet.
        /// </param>
        /// <param name="encoding">Text encoding during output text in CSV format.</param>
        public void ExportAsCSV(string path, int startRow = 0, Encoding encoding = null)
        {
            ExportAsCSV(path, new RangePosition(startRow, 0, -1, -1), encoding);
        }

        /// <summary>
        ///     Export spreadsheet as CSV format from specified range.
        /// </summary>
        /// <param name="path">File path to write CSV format as stream.</param>
        /// <param name="addressOrName">Range to be output from this worksheet, specified by address or name.</param>
        /// <param name="encoding">Text encoding during output text in CSV format.</param>
        public void ExportAsCSV(string path, string addressOrName, Encoding encoding = null)
        {
            if (RangePosition.IsValidAddress(addressOrName))
            {
                ExportAsCSV(path, new RangePosition(addressOrName), encoding);
            }
            else
            {
                NamedRange namedRange;
                if (TryGetNamedRange(addressOrName, out namedRange))
                    ExportAsCSV(path, namedRange, encoding);
                else
                    throw new InvalidAddressException(addressOrName);
            }
        }

        /// <summary>
        ///     Export spreadsheet as CSV format from specified range.
        /// </summary>
        /// <param name="path">File path to write CSV format as stream.</param>
        /// <param name="range">Range to be output from this worksheet.</param>
        /// <param name="encoding">Text encoding during output text in CSV format.</param>
        public void ExportAsCSV(string path, RangePosition range, Encoding encoding = null)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                ExportAsCSV(fs, range, encoding);
            }
        }

        /// <summary>
        ///     Export spreadsheet as CSV format from specified number of rows.
        /// </summary>
        /// <param name="s">Stream to write CSV format as stream.</param>
        /// <param name="startRow">
        ///     Number of rows start to export data,
        ///     this property is useful to skip the headers on top of worksheet.
        /// </param>
        /// <param name="encoding">Text encoding during output text in CSV format</param>
        public void ExportAsCSV(Stream s, int startRow = 0, Encoding encoding = null)
        {
            ExportAsCSV(s, new RangePosition(startRow, 0, -1, -1), encoding);
        }

        public string ExportAsTxt()
        {
            var range = FixRange(new RangePosition(0, 0, -1, -1));
            var maxRow = Math.Min(range.EndRow, MaxContentRow);
            var maxCol = Math.Min(range.EndCol, MaxContentCol);
            var sb = new StringBuilder();
            for (var r = range.Row; r <= maxRow; r++)
            {
                for (var c = range.Col; c <= maxCol;)
                {
                    var cell = GetCell(r, c);
                    if (cell == null || !cell.IsValidCell)
                    {
                        sb.Append('\t');
                        c++;
                    }
                    else
                    {
                        var data = cell.Data.ToString();
                        sb.Append(data);
                        sb.Append('\t');
                        c += cell.Colspan;
                    }
                }

                sb.Append('\n');
            }

            return sb.ToString();
        }

        public void ExportAsTxt(Stream s, int startRow = 0, Encoding encoding = null)
        {
            ExportAsTxt(s, new RangePosition(startRow, 0, -1, -1), encoding);
        }

        public void ExportAsTxt(Stream s, RangePosition range, Encoding encoding = null)
        {
            range = FixRange(range);

            var maxRow = Math.Min(range.EndRow, MaxContentRow);
            var maxCol = Math.Min(range.EndCol, MaxContentCol);

            if (encoding == null) encoding = Encoding.Default;

            using (var sw = new StreamWriter(s, encoding))
            {
                var sb = new StringBuilder();

                for (var r = range.Row; r <= maxRow; r++)
                {
                    if (sb.Length > 0)
                    {
                        sw.WriteLine(sb.ToString());
                        sb.Length = 0;
                    }

                    for (var c = range.Col; c <= maxCol;)
                    {
                        if (sb.Length > 0) sb.Append('\t');

                        var cell = GetCell(r, c);
                        if (cell == null || !cell.IsValidCell)
                        {
                            c++;
                        }
                        else
                        {
                            var data = cell.Data;

                            var quota = false;
                            //if (!quota)
                            //{
                            //	if (cell.DataFormat == CellDataFormatFlag.Text)
                            //	{
                            //		quota = true;
                            //	}
                            //}

                            var str = data as string;
                            if (str != null)
                            {
                                if (!string.IsNullOrEmpty(str)
                                    && (cell.DataFormat == CellDataFormatFlag.Text
                                        || str.IndexOf(',') >= 0 || str.IndexOf('"') >= 0
                                        || str.StartsWith(" ") || str.EndsWith(" ")))
                                    quota = true;
                            }
                            else
                            {
                                str = Convert.ToString(data);
                            }

                            if (quota)
                            {
                                sb.Append('"');
                                sb.Append(str.Replace("\"", "\"\""));
                                sb.Append('"');
                            }
                            else
                            {
                                sb.Append(str);
                            }

                            c += cell.Colspan;
                        }
                    }
                }

                if (sb.Length > 0)
                {
                    sw.WriteLine(sb.ToString());
                    sb.Length = 0;
                }
            }
        }

        /// <summary>
        ///     Export spreadsheet as CSV format from specified range.
        /// </summary>
        /// <param name="s">Stream to write CSV format as stream.</param>
        /// <param name="addressOrName">Range to be output from this worksheet, specified by address or name.</param>
        /// <param name="encoding">Text encoding during output text in CSV format.</param>
        public void ExportAsCSV(Stream s, string addressOrName, Encoding encoding = null)
        {
            if (RangePosition.IsValidAddress(addressOrName))
            {
                ExportAsCSV(s, new RangePosition(addressOrName), encoding);
            }
            else
            {
                NamedRange namedRange;
                if (TryGetNamedRange(addressOrName, out namedRange))
                    ExportAsCSV(s, namedRange, encoding);
                else
                    throw new InvalidAddressException(addressOrName);
            }
        }

        /// <summary>
        ///     Export spreadsheet as CSV format from specified range.
        /// </summary>
        /// <param name="s">Stream to write CSV format as stream.</param>
        /// <param name="range">Range to be output from this worksheet.</param>
        /// <param name="encoding">Text encoding during output text in CSV format.</param>
        public void ExportAsCSV(Stream s, RangePosition range, Encoding encoding = null)
        {
            range = FixRange(range);

            var maxRow = Math.Min(range.EndRow, MaxContentRow);
            var maxCol = Math.Min(range.EndCol, MaxContentCol);

            if (encoding == null) encoding = Encoding.Default;

            using (var sw = new StreamWriter(s, encoding))
            {
                var sb = new StringBuilder();

                for (var r = range.Row; r <= maxRow; r++)
                {
                    if (sb.Length > 0)
                    {
                        sw.WriteLine(sb.ToString());
                        sb.Length = 0;
                    }

                    for (var c = range.Col; c <= maxCol;)
                    {
                        if (sb.Length > 0) sb.Append(',');

                        var cell = GetCell(r, c);
                        if (cell == null || !cell.IsValidCell)
                        {
                            c++;
                        }
                        else
                        {
                            var data = cell.Data;

                            var quota = false;
                            //if (!quota)
                            //{
                            //	if (cell.DataFormat == CellDataFormatFlag.Text)
                            //	{
                            //		quota = true;
                            //	}
                            //}

                            var str = data as string;
                            if (str != null)
                            {
                                if (!string.IsNullOrEmpty(str)
                                    && (cell.DataFormat == CellDataFormatFlag.Text
                                        || str.IndexOf(',') >= 0 || str.IndexOf('"') >= 0
                                        || str.StartsWith(" ") || str.EndsWith(" ")))
                                    quota = true;
                            }
                            else
                            {
                                str = Convert.ToString(data);
                            }

                            if (quota)
                            {
                                sb.Append('"');
                                sb.Append(str.Replace("\"", "\"\""));
                                sb.Append('"');
                            }
                            else
                            {
                                sb.Append(str);
                            }

                            c += cell.Colspan;
                        }
                    }
                }

                if (sb.Length > 0)
                {
                    sw.WriteLine(sb.ToString());
                    sb.Length = 0;
                }
            }
        }

        #endregion // Export
    }
}