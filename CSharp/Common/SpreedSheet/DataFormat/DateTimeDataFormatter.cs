#define WPF

using System;
using System.Globalization;
using System.Threading;
using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Utility;

namespace unvell.ReoGrid.DataFormat
{
    /// <summary>
    ///     Datetime data formatter
    /// </summary>
    public class DateTimeDataFormatter : IDataFormatter
    {
        /// <summary>
        ///     Base start time used to calculcate the date from a number value
        /// </summary>
        public static DateTime BaseStartDate { get; set; } = new DateTime(1900, 1, 1);

        /// <summary>
        ///     Format cell
        /// </summary>
        /// <param name="cell">cell to be formatted</param>
        /// <returns>Formatted text used to display as cell content</returns>
        public string FormatCell(Cell cell)
        {
            var data = cell.InnerData;

            var isFormat = false;
            double number;
            var value = BaseStartDate;
            string formattedText = null;

            if (data is DateTime)
            {
                value = (DateTime)data;
                isFormat = true;
            }
            else if (CellUtility.TryGetNumberData(data, out number))
            {
                try
                {
                    // Excel/Lotus 2/29/1900 bug   
                    // original post: http://stackoverflow.com/questions/4538321/reading-datetime-value-from-excel-sheet
                    value = DateTime.FromOADate(number);

                    isFormat = true;
                }
                catch
                {
                }
            }
            else
            {
                var strdata = data is string ? (string)data : Convert.ToString(data);

                double days = 0;
                if (double.TryParse(strdata, out days))
                    try
                    {
                        value = value.AddDays(days);
                        isFormat = true;
                    }
                    catch
                    {
                    }
                else
                    isFormat = DateTime.TryParse(strdata, out value);
            }

            if (isFormat)
            {
                if (cell.InnerStyle.HAlign == GridHorAlign.General) cell.RenderHorAlign = GridRenderHorAlign.Right;

                CultureInfo culture = null;

                var pattern = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;

                if (cell.DataFormatArgs != null && cell.DataFormatArgs is DateTimeFormatArgs)
                {
                    var dargs = (DateTimeFormatArgs)cell.DataFormatArgs;

                    // fixes issue #203: pattern is ignored incorrectly
                    if (!string.IsNullOrEmpty(dargs.Format)) pattern = dargs.Format;

                    culture = dargs.CultureName == null
                              || string.Equals(dargs.CultureName, Thread.CurrentThread.CurrentCulture.Name)
                        ? Thread.CurrentThread.CurrentCulture
                        : new CultureInfo(dargs.CultureName);
                }
                else
                {
                    culture = Thread.CurrentThread.CurrentCulture;
                    cell.DataFormatArgs = new DateTimeFormatArgs { Format = pattern, CultureName = culture.Name };
                }

                if (culture.Name.StartsWith("ja") && pattern.Contains("g"))
                {
                    culture = new CultureInfo("ja-JP", true);
                    culture.DateTimeFormat.Calendar = new JapaneseCalendar();
                }

                try
                {
                    switch (pattern)
                    {
                        case "d":
                            formattedText = value.Day.ToString();
                            break;

                        default:
                            formattedText = value.ToString(pattern, culture);
                            break;
                    }
                }
                catch
                {
                    formattedText = Convert.ToString(value);
                }
            }

            return isFormat ? formattedText : null;
        }

        /// <summary>
        ///     Determines whether or not to perform a test when target cell is not set as datetime format.
        /// </summary>
        /// <returns></returns>
        public bool PerformTestFormat()
        {
            return true;
        }

        /// <summary>
        ///     Represents the argument that is used during format a cell as data time.
        /// </summary>
        [Serializable]
        public struct DateTimeFormatArgs
        {
            private string format;

            /// <summary>
            ///     Get or set the date time pattern. (Standard .NET datetime pattern is supported, e.g.: yyyy/MM/dd)
            /// </summary>
            public string Format
            {
                get { return format; }
                set { format = value; }
            }

            private string cultureName;

            /// <summary>
            ///     Get or set the culture name that is used to format datetime according to localization settings.
            /// </summary>
            public string CultureName
            {
                get { return cultureName; }
                set { cultureName = value; }
            }

            /// <summary>
            ///     Compare to another object, check whether or not two objects are same.
            /// </summary>
            /// <param name="obj">Another object to be compared.</param>
            /// <returns>True if two objects are same; Otherwise return false.</returns>
            public override bool Equals(object obj)
            {
                if (!(obj is DateTimeFormatArgs)) return false;
                var o = (DateTimeFormatArgs)obj;
                return format.Equals(o.format)
                       && cultureName.Equals(o.cultureName);
            }

            /// <summary>
            ///     Get the hash code of this argument object.
            /// </summary>
            /// <returns>Hash code of argument object.</returns>
            public override int GetHashCode()
            {
                return format.GetHashCode() ^ cultureName.GetHashCode();
            }
        }
    }
}