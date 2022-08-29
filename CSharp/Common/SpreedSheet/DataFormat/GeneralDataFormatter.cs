#define WPF

using System;
using SpreedSheet.Core.Enum;

namespace unvell.ReoGrid.DataFormat
{
    /// <summary>
    ///     GeneralDataFormatter supports both Text and Numeric format.
    ///     And format type can be switched after data changed by user inputing.
    /// </summary>
    internal class GeneralDataFormatter : IDataFormatter
    {
        public string[] Formats
        {
            get { return null; }
        }

        public string FormatCell(Cell cell)
        {
            var data = cell.InnerData;

            // check numeric
            var isNumeric = false;

            double value = 0;
            if (data is int)
            {
                value = (int)data;
                isNumeric = true;
            }
            else if (data is double)
            {
                value = (double)data;
                isNumeric = true;
            }
            else if (data is float)
            {
                value = (float)data;
                isNumeric = true;
            }
            else if (data is long)
            {
                value = (long)data;
                isNumeric = true;
            }
            else if (data is short)
            {
                value = (short)data;
                isNumeric = true;
            }
            else if (data is decimal)
            {
                value = (double)(decimal)data;
                isNumeric = true;
            }
            else if (data is string)
            {
                var str = (string)data;

                if (str.StartsWith(" ") || str.EndsWith(" ")) str = str.Trim();

                isNumeric = double.TryParse(str, out value);

                if (isNumeric) cell.InnerData = value;
            }

            if (isNumeric)
            {
                if (cell.InnerStyle.HAlign == GridHorAlign.General) cell.RenderHorAlign = GridRenderHorAlign.Right;

                return Convert.ToString(value);
            }

            return null;
        }

        public bool PerformTestFormat()
        {
            return true;
        }
    }
}