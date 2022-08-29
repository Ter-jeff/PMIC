#define WPF

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using SpreedSheet.Core;
using SpreedSheet.Core.Workbook;

namespace unvell.ReoGrid.IO
{
    internal static class CSVFormat
    {
        public const int DEFAULT_READ_BUFFER_LINES = 512;

        private static readonly Regex lineRegex =
            new Regex("\\s*(\\\"(?<item>[^\\\"]*)\\\"|(?<item>[^,]*))\\s*,?", RegexOptions.Compiled);

        public static void Read(Stream stream, Worksheet sheet, RangePosition targetRange,
            Encoding encoding = null, int bufferLines = DEFAULT_READ_BUFFER_LINES, bool autoSpread = true)
        {
            targetRange = sheet.FixRange(targetRange);

            var lines = new string[bufferLines];
            var bufferLineList = new List<object>[bufferLines];

            for (var i = 0; i < bufferLineList.Length; i++) bufferLineList[i] = new List<object>(256);

#if DEBUG
            var sw = Stopwatch.StartNew();
#endif

            var row = targetRange.Row;
            var totalReadLines = 0;

            using (var sr = new StreamReader(stream, encoding))
            {
                sheet.SuspendDataChangedEvents();
                var maxCols = 0;

                try
                {
                    var finished = false;
                    while (!finished)
                    {
                        var readLines = 0;

                        for (; readLines < lines.Length; readLines++)
                        {
                            var line = sr.ReadLine();
                            if (line == null)
                            {
                                finished = true;
                                break;
                            }

                            lines[readLines] = line;

                            totalReadLines++;
                            if (!autoSpread && totalReadLines > targetRange.Rows)
                            {
                                finished = true;
                                break;
                            }
                        }

                        if (autoSpread && row + readLines > sheet.RowCount)
                        {
                            var appendRows = bufferLines - sheet.RowCount % bufferLines;
                            if (appendRows <= 0) appendRows = bufferLines;
                            sheet.AppendRows(appendRows);
                        }

                        for (var i = 0; i < readLines; i++)
                        {
                            var line = lines[i];

                            var toBuffer = bufferLineList[i];
                            toBuffer.Clear();

                            foreach (Match m in lineRegex.Matches(line))
                            {
                                toBuffer.Add(m.Groups["item"].Value);

                                if (toBuffer.Count >= targetRange.Cols) break;
                            }

                            if (maxCols < toBuffer.Count) maxCols = toBuffer.Count;

                            if (autoSpread && maxCols >= sheet.ColumnCount) sheet.SetCols(maxCols + 1);
                        }

                        sheet.SetRangeData(row, targetRange.Col, readLines, maxCols, bufferLineList);
                        row += readLines;
                    }
                }
                finally
                {
                    sheet.ResumeDataChangedEvents();
                }

                sheet.RaiseRangeDataChangedEvent(new RangePosition(
                    targetRange.Row, targetRange.Col, maxCols, totalReadLines));
            }

#if DEBUG
            sw.Stop();
            Debug.WriteLine("load csv file: " + sw.ElapsedMilliseconds + " ms, rows: " + row);
#endif
        }
    }

    #region CSV File Provider

    internal class CSVFileFormatProvider : IFileFormatProvider
    {
        public bool IsValidFormat(string file)
        {
            return Path.GetExtension(file).Equals(".csv", StringComparison.CurrentCultureIgnoreCase);
        }

        public void Load(Workbook workbook, Stream stream, Encoding encoding, object arg, string sheetName)
        {
            var autoSpread = true;
            var bufferLines = CSVFormat.DEFAULT_READ_BUFFER_LINES;
            var targetRange = RangePosition.EntireRange;

            var csvArg = arg as CSVFormatArgument;

            if (csvArg != null)
            {
                autoSpread = csvArg.AutoSpread;
                bufferLines = csvArg.BufferLines;
                targetRange = csvArg.TargetRange;
            }

            Worksheet sheet = null;

            if (workbook.Worksheets.Count == 0)
            {
                sheet = workbook.CreateWorksheet("Sheet1");
                workbook.Worksheets.Add(sheet);
            }
            else
            {
                while (workbook.Worksheets.Count > 1) workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);

                sheet = workbook.Worksheets[0];
                sheet.Reset();
            }

            CSVFormat.Read(stream, sheet, targetRange, encoding, bufferLines, autoSpread);
        }

        public void Save(IWorkbook workbook, Stream stream, Encoding encoding, object arg)
        {
            throw new NotSupportedException(
                "Saving entire workbook as CSV is not supported, use Worksheet.ExportAsCSV instead.");
            //int fromRow = 0, fromCol = 0, toRow = 0, toCol = 0;

            //if (args != null)
            //{
            //	object arg;
            //	if (args.TryGetValue("fromRow", out arg)) fromRow = (int)arg;
            //	if (args.TryGetValue("fromCol", out arg)) fromCol = (int)arg;
            //	if (args.TryGetValue("toRow", out arg)) toRow = (int)arg;
            //	if (args.TryGetValue("toCol", out arg)) toCol = (int)arg;
            //}
        }

        public bool IsValidFormat(Stream s)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>
    ///     Arguments for loading and saving CSV format.
    /// </summary>
    public class CSVFormatArgument
    {
        /// <summary>
        ///     Create the argument object instance.
        /// </summary>
        public CSVFormatArgument()
        {
            AutoSpread = true;
            BufferLines = CSVFormat.DEFAULT_READ_BUFFER_LINES;
            SheetName = "Sheet1";
            TargetRange = RangePosition.EntireRange;
        }

        /// <summary>
        ///     Determines whether or not allow to expand worksheet to load more data from CSV file. (Default is True)
        /// </summary>
        public bool AutoSpread { get; set; }

        /// <summary>
        ///     Determines how many rows read from CSV file one time. (Default is CSVFormat.DEFAULT_READ_BUFFER_LINES = 512)
        /// </summary>
        public int BufferLines { get; set; }

        /// <summary>
        ///     Determines the default worksheet name if CSV file loaded into a new workbook. (Default is "Sheet1")
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        ///     Determines where to start import CSV data on worksheet.
        /// </summary>
        public RangePosition TargetRange { get; set; }
    }

    #endregion // CSV File Provider
}