#define WPF

using System;
using System.IO;
using System.Text;
using SpreedSheet.Core.Workbook;
using unvell.ReoGrid.IO.OpenXML;

namespace unvell.ReoGrid.IO
{
    /// <summary>
    ///     File format flag
    /// </summary>
    public enum FileFormat : ushort
    {
        /// <summary>
        ///     ReoGrid Format (RGF)
        /// </summary>
        ReoGridFormat = 1,

        /// <summary>
        ///     CSV Plain-text format
        /// </summary>
        CSV = 10,

        /// <summary>
        ///     Excel 2007 (Office OpenXML-based format)
        /// </summary>
        Excel2007 = 30,

        /// <summary>
        ///     Decide file format by extension automatically
        /// </summary>
        _Auto = 0,

        /// <summary>
        ///     User-defined file format provider (Reserved)
        /// </summary>
        _Custom = 100,
        IGXL = 101
    }

    #region File Format Provider Interface

    /// <summary>
    ///     Interface of file format provider
    /// </summary>
    internal interface IFileFormatProvider
    {
        /// <summary>
        ///     Check whether specified filepath is valid file format to be processed.
        /// </summary>
        /// <param name="file">file path</param>
        /// <returns>true if specified file is valid format</returns>
        bool IsValidFormat(string file);

        /// <summary>
        ///     Load spreadsheet from specified stream
        /// </summary>
        /// <param name="workbook">ReoGrid workbook to be loaded</param>
        /// <param name="stream">Stream to input serialized data of workbook</param>
        /// <param name="encoding">Encoding used to read plain-text file format</param>
        /// <param name="arg">Arguments of format provider</param>
        /// <param name="sheetName"></param>
        void Load(Workbook workbook, Stream stream, Encoding encoding, object arg, string sheetName);

        /// <summary>
        ///     Save spreadsheet to specified stream
        /// </summary>
        /// <param name="workbook">ReoGrid workbook to be saved</param>
        /// <param name="stream">Stream to output serialized data of workbook</param>
        /// <param name="encoding">Encoding used to write plain-text file format</param>
        /// <param name="arg">Arguments of format provider</param>
        void Save(IWorkbook workbook, Stream stream, Encoding encoding, object arg);
    }

    #endregion // File Format Provider Interface

    #region RGF Provider

    /// <summary>
    ///     Represents the file format provider for saving and loading workbook and worksheets
    /// </summary>
    internal class ReoGridFileFormatProvider : IFileFormatProvider
    {
        /// <summary>
        ///     Check whether or not the file is valid format of this provider
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public bool IsValidFormat(string file)
        {
            return Path.GetExtension(file).Equals(".rgf", StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        ///     Load workbook from specified input stream
        /// </summary>
        /// <param name="workbook">Workbook to be loaded</param>
        /// <param name="stream">Input stream</param>
        /// <param name="encoding">Encoding used to read text-based stream, such as XML</param>
        /// <param name="arg">Provider custom parameters</param>
        public void Load(Workbook workbook, Stream stream, Encoding encoding, object arg, string sheetName)
        {
            Worksheet sheet = null;

            if (workbook.Worksheets.Count == 0)
            {
                sheet = workbook.CreateWorksheet("Sheet1");
                workbook.Worksheets.Add(sheet);
            }
            else
            {
                sheet = workbook.Worksheets[0];
            }

            sheet.LoadRGF(stream);
        }

        public void Save(IWorkbook workbook, Stream stream, Encoding encoding, object arg)
        {
            if (workbook.Worksheets == null || workbook.Worksheets.Count <= 0) return;

            workbook.Worksheets[0].Save(stream);
        }
    }

    /// <summary>
    ///     Class that contains some arguments for reading and saving RGF format.
    /// </summary>
    public class ReoGridFormatArgument
    {
    }

    #endregion RGF Provider

    #region Excel File Provider

    internal class ExcelFileFormatProvider : IFileFormatProvider
    {
        public bool IsValidFormat(string file)
        {
            return Path.GetExtension(file).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase);
        }

        public void Load(Workbook workbook, Stream stream, Encoding encoding, object arg, string sheetName)
        {
            ExcelReader.ReadStream(workbook, stream);
        }

        public void Save(IWorkbook workbook, Stream stream, Encoding encoding, object arg)
        {
            ExcelWriter.WriteStream(workbook, stream);
        }
    }

    /// <summary>
    ///     Class that contains some arguments for reading and saving Excel format.
    /// </summary>
    public class ExcelFileFormatArgument
    {
    }

    #endregion
}