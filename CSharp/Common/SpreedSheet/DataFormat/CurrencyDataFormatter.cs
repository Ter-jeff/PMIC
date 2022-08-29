﻿#define WPF

using System;
using SpreedSheet.Core.Enum;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid.DataFormat
{
    /// <summary>
    ///     Currency data formatter
    /// </summary>
    public class CurrencyDataFormatter : IDataFormatter
    {
        /// <summary>
        ///     Format specified cell
        /// </summary>
        /// <param name="cell">cell instance</param>
        /// <returns>true if cell has been formatted</returns>
        public string FormatCell(Cell cell)
        {
            var isFormat = false;

            var data = cell.InnerData;
            var currency = double.NaN;

            if (data is double)
            {
                isFormat = true;
                currency = (double)data;
            }
            else if (data is DateTime)
            {
                currency = (new DateTime(1900, 1, 1) - (DateTime)data).TotalDays;
                isFormat = true;
            }
            else
            {
                var str = Convert.ToString(data).Trim();
                var number = string.Empty;

                if (str.StartsWith("$"))
                {
                    number = str.Substring(1);
                    if (double.TryParse(number, out currency))
                    {
                        isFormat = true;
                        cell.InnerData = currency;
                    }
                }
                else
                {
                    // Stop trying to convert datetime value to currency, #170
                    //
                    //DateTime date = new DateTime(1900, 1, 1);
                    //if (DateTime.TryParse(str, out date))
                    //{
                    //	currency = (date - new DateTime(1900, 1, 1)).TotalDays;
                    //	isFormat = true;
                    //}
                    //else
                    //{
                    isFormat = double.TryParse(str, out currency);
                    //}
                }
            }

            if (isFormat)
            {
                if (cell.InnerStyle.HAlign == GridHorAlign.General) cell.RenderHorAlign = GridRenderHorAlign.Right;

                string prefixSymbol = null, postfixSymbol = null;
                short decimals = 2;
                var negativeStyle = NumberDataFormatter.NumberNegativeStyle.Default;
                string prefix = null;
                string postfix = null;

                if (cell.DataFormatArgs != null && cell.DataFormatArgs is CurrencyFormatArgs)
                {
                    var args = (CurrencyFormatArgs)cell.DataFormatArgs;
                    prefixSymbol = args.PrefixSymbol;
                    postfixSymbol = args.PostfixSymbol;
                    decimals = args.DecimalPlaces;
                    negativeStyle = args.NegativeStyle;
                    prefix = args.CustomNegativePrefix;
                    postfix = args.CustomNegativePostfix;
                }
                //else
                //{
                //	var culture = Thread.CurrentThread.CurrentCulture;

                //	switch (culture.NumberFormat.CurrencyPositivePattern)
                //	{
                //		case 0: prefixSymbol = culture.NumberFormat.CurrencySymbol; postfixSymbol = null; break;
                //		case 1: prefixSymbol = null; postfixSymbol = culture.NumberFormat.CurrencySymbol; break;
                //		case 2: prefixSymbol = " " + culture.NumberFormat.CurrencySymbol; postfixSymbol = null; break;
                //		case 3: prefixSymbol = null; postfixSymbol = " " + culture.NumberFormat.CurrencySymbol; break;
                //	}

                //	cell.DataFormatArgs = new CurrencyFormatArgs { PrefixSymbol = prefixSymbol, PostfixSymbol = postfixSymbol, DecimalPlaces = decimals };
                //}

                if (currency < 0)
                {
                    if ((negativeStyle & NumberDataFormatter.NumberNegativeStyle.Red) ==
                        NumberDataFormatter.NumberNegativeStyle.Red)
                        cell.RenderColor = SolidColor.Red;
                    else
                        cell.RenderColor = SolidColor.Transparent;
                }

                // decimal places
                var decimalPlacePart = new string('0', decimals);

                // number
                var numberPartFormat = prefixSymbol + "#,##0." + decimalPlacePart + postfixSymbol;

                if ((negativeStyle & NumberDataFormatter.NumberNegativeStyle.Brackets) ==
                    NumberDataFormatter.NumberNegativeStyle.Brackets)
                    numberPartFormat = currency < 0 ? "(" + numberPartFormat + ")" : numberPartFormat;
                else if ((negativeStyle & NumberDataFormatter.NumberNegativeStyle.Prefix_Sankaku) ==
                         NumberDataFormatter.NumberNegativeStyle.Prefix_Sankaku)
                    numberPartFormat = currency < 0 ? "▲ " + numberPartFormat : numberPartFormat;
                else if ((negativeStyle & NumberDataFormatter.NumberNegativeStyle.CustomSymbol) ==
                         NumberDataFormatter.NumberNegativeStyle.CustomSymbol)
                    numberPartFormat = currency < 0 ? prefix + numberPartFormat + postfix : numberPartFormat;

                // negative
                if ((negativeStyle & NumberDataFormatter.NumberNegativeStyle.Minus) == 0) currency = Math.Abs(currency);

                return currency.ToString(numberPartFormat);
            }

            return null;
        }

        /// <summary>
        ///     Determine whether or not to perform format test
        /// </summary>
        /// <returns>True to perform test; False to abort</returns>
        public bool PerformTestFormat()
        {
            return true;
        }

        /// <summary>
        ///     Represents arguments of currency data format.
        /// </summary>
        [Serializable]
        public class CurrencyFormatArgs : NumberDataFormatter.NumberFormatArgs
        {
            /// <summary>
            ///     Currency symbol that displayed before currency number.
            /// </summary>
            public string PrefixSymbol { get; set; }

            /// <summary>
            ///     Currency symbol that displayed after currency number.
            /// </summary>
            public string PostfixSymbol { get; set; }

            /// <summary>
            ///     Culture name in English. (e.g. en-US)
            /// </summary>
            public string CultureEnglishName { get; set; }

            /// <summary>
            ///     Check whether or not two objects are same.
            /// </summary>
            /// <param name="obj">Another object to be compared.</param>
            /// <returns>True if two objects are same; Otherwise return false.</returns>
            public override bool Equals(object obj)
            {
                if (!(obj is CurrencyFormatArgs)) return false;

                var o = (CurrencyFormatArgs)obj;

                return PrefixSymbol == o.PrefixSymbol
                       && PostfixSymbol == o.PostfixSymbol
                       && string.Compare(CultureEnglishName, o.CultureEnglishName, true) == 0
                       && base.Equals(obj);
            }

            /// <summary>
            ///     Get hash code
            /// </summary>
            /// <returns></returns>
            public override int GetHashCode()
            {
                return base.GetHashCode();
            }
        }
    }
}