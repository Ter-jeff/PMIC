using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using SpreedSheet.Core;
using SpreedSheet.Core.Workbook;
using unvell.ReoGrid;
using unvell.ReoGrid.IO;

namespace SpreedSheet.IO
{
    internal static class TxtFormat
    {
        public const int DEFAULT_READ_BUFFER_LINES = 512;

        public static void Read(Stream stream, Worksheet sheet, RangePosition targetRange,
            Encoding encoding = null, int bufferLines = DEFAULT_READ_BUFFER_LINES, bool autoSpread = true)
        {
            targetRange = sheet.FixRange(targetRange);

            var lines = new string[bufferLines];
            var bufferLineList = new List<string>[bufferLines];

            for (var i = 0; i < bufferLineList.Length; i++) bufferLineList[i] = new List<string>(256);

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
                            var toBuffer = line.Split('\t').ToList();
                            bufferLineList[i] = toBuffer;

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

    internal class TxtFileFormatProvider : IFileFormatProvider
    {
        public bool IsValidFormat(string file)
        {
            return Path.GetExtension(file).Equals(".txt", StringComparison.CurrentCultureIgnoreCase);
        }

        public void Load(Workbook workbook, Stream stream, Encoding encoding, object arg, string sheetName)
        {
            var autoSpread = true;
            var bufferLines = TxtFormat.DEFAULT_READ_BUFFER_LINES;
            var targetRange = RangePosition.EntireRange;

            var csvArg = arg as TxtFormatArgument;

            if (csvArg != null)
            {
                autoSpread = csvArg.AutoSpread;
                bufferLines = csvArg.BufferLines;
                targetRange = csvArg.TargetRange;
            }

            var sheet = workbook.Worksheets[sheetName];
            if (workbook.Worksheets.Count == 1 && workbook.Worksheets.First().Name == "Sheet1")
            {
                sheet = workbook.Worksheets[0];
                sheet.Name = sheetName;
                sheet.Reset();
            }
            else if (sheet == null)
            {
                sheet = workbook.CreateWorksheet(sheetName);
                //workbook.Worksheets.Add(sheet);
                workbook.Worksheets.Insert(0, sheet);
                workbook.ControlInstance.ActiveWorksheet = sheet;
                workbook.controlAdapter.ControlInstance.ActiveWorksheet = sheet;
            }
            else
            {
                workbook.MoveWorksheet(sheet, 0);
                workbook.ControlInstance.ActiveWorksheet = sheet;
                workbook.controlAdapter.ControlInstance.ActiveWorksheet = sheet;
                sheet.Reset();
            }

            TxtFormat.Read(stream, sheet, targetRange, encoding, bufferLines, autoSpread);
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
    public class TxtFormatArgument
    {
        /// <summary>
        ///     Create the argument object instance.
        /// </summary>
        public TxtFormatArgument()
        {
            AutoSpread = true;
            BufferLines = TxtFormat.DEFAULT_READ_BUFFER_LINES;
            SheetName = "Sheet1";
            TargetRange = RangePosition.EntireRange;
        }

        /// <summary>
        ///     Determines whether or not allow to expand worksheet to load more data from CSV file. (Default is True)
        /// </summary>
        public bool AutoSpread { get; set; }

        /// <summary>
        ///     Determines how many rows read from CSV file one time. (Default is TxtFormat.DEFAULT_READ_BUFFER_LINES = 512)
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