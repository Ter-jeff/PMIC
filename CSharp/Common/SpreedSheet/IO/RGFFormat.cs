﻿#define WPF

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Media.Imaging;
using System.Xml.Serialization;
using SpreedSheet.CellTypes;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO;
using unvell.ReoGrid.Utility;
using unvell.ReoGrid.XML;
#if OUTLINE
using unvell.ReoGrid.Outline;
#endif // OUTLINE

namespace unvell.ReoGrid
{
    public interface RGFCustomBodyHandler
    {
        string SaveData(Cell cell);

        object LoadData(Cell cell, string data);
    }

    public static class RGFPersistenceProvider
    {
        internal static Dictionary<Type, string> CustomBodyTypeIdentifiers { get; } = new Dictionary<Type, string>();

        internal static Dictionary<string, RGFCustomBodyHandler> CustomBodyTypeHandlers { get; } =
            new Dictionary<string, RGFCustomBodyHandler>();

        public static CustomBodyTypeProviderCollection CustomBodyTypeProviders { get; } =
            new CustomBodyTypeProviderCollection();
    }

    public class CustomBodyTypeProviderCollection
    {
        internal CustomBodyTypeProviderCollection()
        {
        }

        public void Add(Type type, string identifier, RGFCustomBodyHandler handler)
        {
            RGFPersistenceProvider.CustomBodyTypeIdentifiers[type] = identifier;
            RGFPersistenceProvider.CustomBodyTypeHandlers[identifier] = handler;
        }

        public void Remove(Type type)
        {
            string identifer;
            if (RGFPersistenceProvider.CustomBodyTypeIdentifiers.TryGetValue(type, out identifer))
                RGFPersistenceProvider.CustomBodyTypeHandlers.Remove(identifer);

            RGFPersistenceProvider.CustomBodyTypeIdentifiers.Remove(type);
        }
    }

    partial class Worksheet
    {
        #region Load

        /// <summary>
        ///     Load grid from specified file.
        /// </summary>
        /// <param name="file">Path of file to load grid.</param>
        /// <exception cref="ReoGridLoadException">Exception will be thrown if any errors happened during loading process.</exception>
        public void Load(string file)
        {
            Load(file, Encoding.Default);
        }

        /// <summary>
        ///     Load worksheet from specified input stream.
        /// </summary>
        /// <param name="file">Path of specified file to load worksheet.</param>
        /// <param name="encoding">Encoding used to decode plain-text if need.</param>
        /// <exception cref="ReoGridLoadException">Exception will be thrown if any errors happen during loading.</exception>
        public void Load(string file, Encoding encoding)
        {
            if (file.EndsWith(".csv", StringComparison.CurrentCultureIgnoreCase))
                LoadCSV(file, encoding);
            else if (file.EndsWith(".xlsx", StringComparison.CurrentCultureIgnoreCase))
                throw new NotSupportedException(
                    "Cannot load Excel file into single worksheet, try use Load method of control.");
            else if (file.EndsWith(".xls", StringComparison.CurrentCultureIgnoreCase))
                throw new NotSupportedException("Loading Excel 2003 format is not supported.");
            else
                //if (file.EndsWith(".rgf", StringComparison.CurrentCultureIgnoreCase))
                LoadRGF(file);

            var newName = Path.GetFileNameWithoutExtension(file);
            if (workbook == null)
            {
                name = newName;
            }
            else
            {
                if (!workbook.CheckWorksheetName(newName)) name = workbook.GetAvailableWorksheetName();
            }

            // raise file loading event
            FileLoaded?.Invoke(this, new FileLoadedEventArgs(file));
        }

        /// <summary>
        ///     Load worksheet from specified input stream.
        /// </summary>
        /// <param name="path">Path of file to load worksheet.</param>
        public void LoadRGF(string path)
        {
            using (var s = File.OpenRead(path))
            {
                LoadRGF(s);
            }
        }

        /// <summary>
        ///     Load worksheet from specified input stream.
        /// </summary>
        /// <param name="s">Input stream to read worksheet.</param>
        /// <returns>True if spreadsheet is loaded successfully.</returns>
        /// <exception cref="ReoGridLoadException">Exception will be thrown if any errors happen during loading.</exception>
        public void LoadRGF(Stream s)
        {
#if DEBUG
            var stop = Stopwatch.StartNew();
#endif // DEBUG

            var xmlReader = new XmlSerializer(typeof(RGXmlSheet));
            RGXmlSheet xmlSheet;

            try
            {
                xmlSheet = xmlReader.Deserialize(s) as RGXmlSheet;
            }
            catch (Exception ex)
            {
                throw new ReoGridLoadException("Read xml format error: " + ex.Message, ex);
            }

            Clear();

            var controlSettings = WorksheetSettings.Default;

            ScaleFactor = 1f;

            // todo: clear current view controller and viewport position
            var culture = Thread.CurrentThread.CurrentCulture;

            #region Head

            // head
            if (xmlSheet.head != null)
            {
                var head = xmlSheet.head;

                #region Settings

                if (head.settings != null)
                {
                    var settings = head.settings;

                    if (settings.showGrid != null && !TextFormatHelper.IsSwitchOn(settings.showGrid))
                        controlSettings = controlSettings.Remove(WorksheetSettings.View_ShowGridLine);

                    if (TextFormatHelper.IsSwitchOn(settings.showPageBreakes))
                        controlSettings = controlSettings.Add(WorksheetSettings.View_ShowPageBreaks);
                    else
                        controlSettings = controlSettings.Remove(WorksheetSettings.View_ShowPageBreaks);

                    if (settings.showRowHeader != null && !TextFormatHelper.IsSwitchOn(settings.showRowHeader))
                        controlSettings = controlSettings.Remove(WorksheetSettings.View_ShowRowHeader);

                    if (settings.showColHeader != null && !TextFormatHelper.IsSwitchOn(settings.showColHeader))
                        controlSettings = controlSettings.Remove(WorksheetSettings.View_ShowColumnHeader);
                }

                #endregion // Settings

                var cultureName = head.meta == null ? null : head.meta.culture;

//#pragma warning disable 618
//				// back-forward compatible
//				if (string.IsNullOrEmpty(cultureName)) cultureName = head.culture;
//#pragma warning restore 618

                // apply culture
                if (!string.IsNullOrEmpty(cultureName))
                    try
                    {
                        culture = new CultureInfo(cultureName);
                    }
                    catch (Exception)
                    {
                        Logger.Log("load", "warning: unsupported culture: " + cultureName);
                    }

                // load default header size
                defaultRowHeight = head.defaultRowHeight;
                DefaultColumnWidth = head.defaultColumnWidth;

                // load row header panel width
                if (head.rowHeaderWidth != null)
                {
                    rowHeaderWidth = (ushort)TextFormatHelper.GetPixelValue(head.rowHeaderWidth, 40);
                    _userRowHeaderWidth = true;
                }

                // selection mode
                if (!string.IsNullOrEmpty(head.selectionMode))
                    selectionMode = XmlFileFormatHelper.DecodeSelectionMode(head.selectionMode);

                // selection style
                if (!string.IsNullOrEmpty(head.selectionStyle))
                    selectionStyle = XmlFileFormatHelper.DecodeSelectionStyle(head.selectionStyle);

                // focus forward direction
                if (!string.IsNullOrEmpty(head.focusForwardDirection))
                    selectionForwardDirection =
                        XmlFileFormatHelper.DecodeFocusForwardDirection(head.focusForwardDirection);

                #region Print

#if PRINT
				// print settings
				if (head.printSettings != null)
				{
					var ps = head.printSettings;

					if (this.printSettings == null)
					{
						this.printSettings = new Print.PrintSettings();
					}

					int[] pageBreakCols = null, pageBreakRows = null;

					if (!string.IsNullOrEmpty(ps.paperName))
					{
						this.printSettings.PaperName = ps.paperName;

					}
					else
					{
						this.printSettings.PaperName = unvell.ReoGrid.Print.PaperSize.Custom.ToString();
					}

					if (!string.IsNullOrEmpty(ps.paperWidth))
					{
						this.printSettings.PaperWidth = TextFormatHelper.GetFloatValue(ps.paperWidth, 8.5f);
					}
					if (!string.IsNullOrEmpty(ps.paperHeight))
					{
						this.printSettings.PaperHeight = TextFormatHelper.GetFloatValue(ps.paperHeight, 11f);
					}

					if (TextFormatHelper.IsSwitchOn(ps.landscape))
					{
						this.printSettings.Landscape = true;
					}

					// load page breaks
					if (!string.IsNullOrEmpty(ps.pageBreakCols))
					{
						pageBreakCols = TextFormatHelper.DecodeIntArray(ps.pageBreakCols);
					}

					if (!string.IsNullOrEmpty(ps.pageBreakRows))
					{
						pageBreakRows = TextFormatHelper.DecodeIntArray(ps.pageBreakRows);
					}

					float pagingScaling = 1f;

					if (!string.IsNullOrEmpty(ps.scaling)
						&& float.TryParse(ps.scaling, out pagingScaling))
					{
						this.PrintSettings.PageScaling = pagingScaling;
					}

					if (!string.IsNullOrEmpty(ps.pageOrder))
					{
						this.PrintSettings.PageOrder = XmlFileFormatHelper.DecodePageOrder(ps.pageOrder);
					}

#if DEBUG
					Debug.Assert(pageBreakRows == null || pageBreakRows.Length >= 2);
					Debug.Assert(pageBreakCols == null || pageBreakCols.Length >= 2);
#endif // DEBUG

					if (pageBreakCols != null && pageBreakCols.Length >= 2
						&& pageBreakRows != null && pageBreakRows.Length >= 2)
					{
						if (this.userPageBreakCols == null)
							this.userPageBreakCols = new List<int>();
						else
							this.userPageBreakCols.Clear();

						if (this.userPageBreakRows == null)
							this.userPageBreakRows = new List<int>();
						else
							this.userPageBreakRows.Clear();

						this.userPageBreakCols.AddRange(pageBreakCols);
						this.userPageBreakRows.AddRange(pageBreakRows);

						int printRangeRow = pageBreakRows[0];
						int printRangeRow2 = pageBreakRows[pageBreakRows.Length - 1];
						int printRangeCol = pageBreakCols[0];
						int printRangeCol2 = pageBreakCols[pageBreakCols.Length - 1];

						this.printableRange = new RangePosition(printRangeRow, printRangeCol,
							printRangeRow2 - printRangeRow, printRangeCol2 - printRangeCol);

						if (ps.margins != null)
						{
							this.printSettings.Margins = new Print.PageMargins(
								(float)ps.margins.top, (float)ps.margins.bottom,
								(float)ps.margins.left, (float)ps.margins.right);
						}
					}
				}
#endif // PRINT

                #endregion // Print
            }

            #endregion // Head

            #region Root Style

            // root style
            if (xmlSheet.style != null)
            {
                RootStyle = StyleUtility.ConvertFromXmlStyle(this, xmlSheet.style, culture);

                StyleUtility.CopyStyle(DefaultStyle, RootStyle, DefaultStyle.Flag & ~RootStyle.Flag);
            }

            #endregion // Root Style

            // cols and rows
            Resize(xmlSheet.head.rows, xmlSheet.head.cols);

            #region Columns

            foreach (var xmlCol in xmlSheet.cols)
            {
                var colhead = cols[xmlCol.col];

                if (maxColumnHeader < colhead.Index) maxColumnHeader = colhead.Index;

                colhead.InnerWidth = xmlCol.width;

                colhead.LastWidth = (ushort)TextFormatHelper.GetPixelValue(xmlCol.lastWidth, 0);
                colhead.IsAutoWidth = string.IsNullOrEmpty(xmlCol.autoWidth)
                    ? true
                    : TextFormatHelper.IsSwitchOn(xmlCol.autoWidth);

                if (!string.IsNullOrEmpty(xmlCol.text)) colhead.Text = xmlCol.text;

                if (!string.IsNullOrEmpty(xmlCol.textColor))
                {
                    SolidColor textColor;

                    if (TextFormatHelper.DecodeColor(xmlCol.textColor, out textColor)) colhead.TextColor = textColor;
                }

                if (xmlCol.style != null)
                {
                    colhead.InnerStyle = new WorksheetRangeStyle(RootStyle);

                    StyleUtility.CopyStyle(StyleUtility.ConvertFromXmlStyle(
                        this, xmlCol.style, culture), cols[xmlCol.col].InnerStyle);
                }

                if (!string.IsNullOrEmpty(xmlCol.defaultCellBody))
                {
                    Type type = null;

                    if (CellTypesManager.CellTypes.TryGetValue(xmlCol.defaultCellBody, out type))
                        colhead.DefaultCellBody = type;
                }
            }

            #endregion // Columns

            #region Rows

            foreach (var row in xmlSheet.rows)
            {
                var rowhead = rows[row.row];

                if (maxRowHeader < rowhead.Index) maxRowHeader = rowhead.Index;

                rowhead.InnerHeight = row.height;

                rowhead.LastHeight = (ushort)TextFormatHelper.GetPixelValue(row.lastHeight, 0);
                rowhead.IsAutoHeight = string.IsNullOrEmpty(row.autoHeight)
                    ? true
                    : TextFormatHelper.IsSwitchOn(row.autoHeight);

                if (!string.IsNullOrEmpty(row.text)) rowhead.Text = row.text;

                if (!string.IsNullOrEmpty(row.textColor))
                {
                    SolidColor textColor;

                    if (TextFormatHelper.DecodeColor(row.textColor, out textColor)) rowhead.TextColor = textColor;
                }

                if (row.style != null)
                {
                    rowhead.InnerStyle = new WorksheetRangeStyle(RootStyle);

                    StyleUtility.CopyStyle(StyleUtility.ConvertFromXmlStyle(
                        this, row.style, culture), rows[row.row].InnerStyle);
                }
            }

            #endregion // Rows

            #region Normalize Columns and Rows

            var left = cols.Count > 0 ? cols[0].InnerWidth : 0;
            for (var c = 1; c < cols.Count; c++)
            {
                cols[c].Left = left;
                left += cols[c].InnerWidth;
            }

            var top = rows.Count > 0 ? rows[0].InnerHeight : 0;
            for (var r = 1; r < rows.Count; r++)
            {
                rows[r].Top = top;
                top += rows[r].InnerHeight;
            }

            #endregion // Normalize Columns and Rows

            #region Outline

#if OUTLINE
			// load outlines
			if (xmlSheet.head.outlines != null)
			{
				if (xmlSheet.head.outlines.rowOutlines != null)
				{
					foreach (var xmlRowOutline in xmlSheet.head.outlines.rowOutlines)
					{
						var outline = AddOutline(RowOrColumn.Row, xmlRowOutline.start, xmlRowOutline.count);
						if (xmlRowOutline.collapsed) outline.Collapse();
					}
				}

				if (xmlSheet.head.outlines.colOutlines != null)
				{
					foreach (var xmlColOutline in xmlSheet.head.outlines.colOutlines)
					{
						var outline = AddOutline(RowOrColumn.Column, xmlColOutline.start, xmlColOutline.count);
						if (xmlColOutline.collapsed) outline.Collapse();
					}
				}
			}
#endif // OUTLINE

            #endregion // Outline

            #region Named Ranges

            // load named ranges
            if (xmlSheet.head.namedRanges != null)
                foreach (var nr in xmlSheet.head.namedRanges)
                    if (RangePosition.IsValidAddress(nr.address))
                        AddNamedRange(new NamedRange(this, nr.name, new RangePosition(nr.address))
                            { Comment = nr.comment });

            #endregion // Named Ranges

            #region Borders

            foreach (var b in xmlSheet.hborder)
                SetHBorders(b.row, b.col, b.cols, b.StyleGridBorder, XmlFileFormatHelper.DecodeHBorderOwnerPos(b.pos));

            foreach (var b in xmlSheet.vborder)
                SetVBorders(b.row, b.col, b.rows, b.StyleGridBorder, XmlFileFormatHelper.DecodeVBorderOwnerPos(b.pos));

            #endregion // Borders

            List<Cell> cellsTracePrecedents = null;
            List<Cell> cellsTraceDependents = null;
            //List<ReoGridCell> cellsNeedRecalc = null; // TODO: how to recalc dependent cells

            suspendDataChangedEvent = true;

            #region Cells

            foreach (var xmlCell in xmlSheet.cells)
            {
                var rowspan = 1;
                var colspan = 1;
                if (xmlCell.rowspan != null) int.TryParse(xmlCell.rowspan, out rowspan);
                if (xmlCell.colspan != null) int.TryParse(xmlCell.colspan, out colspan);

                var cell = CreateCell(xmlCell.row, xmlCell.col);

                if (rowspan > 1 || colspan > 1)
                    MergeRange(new RangePosition(xmlCell.row, xmlCell.col, rowspan, colspan), updateUIAndEvent: false);

                cell.DataFormat = XmlFileFormatHelper.DecodeCellDataFormat(xmlCell.dataFormat);

                if (xmlCell.style != null)
                {
                    var style = StyleUtility.ConvertFromXmlStyle(this, xmlCell.style, culture);
                    if (style != null) SetCellStyle(cell, style, StyleParentKind.Own);
                }

                #region Data Format

                if (xmlCell.dataFormatArgs != null && cell.DataFormat != CellDataFormatFlag.General)
                {
                    var xmlFormatArgs = xmlCell.dataFormatArgs;

                    object formatArgs = null;

                    switch (cell.DataFormat)
                    {
                        case CellDataFormatFlag.Number:
                            formatArgs = new NumberDataFormatter.NumberFormatArgs
                            {
                                DecimalPlaces =
                                    (short)TextFormatHelper.GetFloatValue(xmlFormatArgs.decimalPlaces, 2, culture),
                                NegativeStyle =
                                    XmlFileFormatHelper.DecodeNegativeNumberStyle(xmlFormatArgs.negativeStyle),
                                UseSeparator = TextFormatHelper.IsSwitchOn(xmlFormatArgs.useSeparator)
                            };
                            break;

                        case CellDataFormatFlag.DateTime:
                            formatArgs = new DateTimeDataFormatter.DateTimeFormatArgs
                            {
                                CultureName = xmlFormatArgs.culture,
                                Format = xmlFormatArgs.pattern
                            };
                            break;

                        case CellDataFormatFlag.Currency:
                        {
                            var cptrn = xmlFormatArgs.pattern;

                            string prefix, postfix;

                            var pindex = cptrn.IndexOf(',');

                            if (pindex > -1)
                            {
                                prefix = cptrn.Substring(0, pindex);
                                postfix = cptrn.Substring(pindex + 1);
                            }
                            else
                            {
                                prefix = cptrn;
                                postfix = null;
                            }

                            formatArgs = new CurrencyDataFormatter.CurrencyFormatArgs
                            {
                                DecimalPlaces =
                                    (short)TextFormatHelper.GetFloatValue(xmlFormatArgs.decimalPlaces, 2, culture),
                                CultureEnglishName = xmlFormatArgs.culture,
                                NegativeStyle =
                                    XmlFileFormatHelper.DecodeNegativeNumberStyle(xmlFormatArgs.negativeStyle),
                                PrefixSymbol = prefix,
                                PostfixSymbol = postfix
                            };
                        }
                            break;

                        case CellDataFormatFlag.Percent:
                            formatArgs = new NumberDataFormatter.NumberFormatArgs
                            {
                                DecimalPlaces =
                                    (short)TextFormatHelper.GetFloatValue(xmlFormatArgs.decimalPlaces, 0, culture)
                            };
                            break;
                    }

                    if (formatArgs != null) cell.DataFormatArgs = formatArgs;
                }

                #endregion // Data Format

                // formula
                string formula = null;
                string cellValue = null;

#if FORMULA
				if (xmlCell.formula != null && !string.IsNullOrEmpty(xmlCell.formula.val))
				{
					formula = xmlCell.formula.val;
					cellValue = xmlCell.data;
				}
				else if (xmlCell.data != null && xmlCell.data.StartsWith("="))
				{
					formula = xmlCell.data.Substring(1);
				}
				else
				{
					cellValue = xmlCell.data;
				}
#else
                cellValue = xmlCell.data;
#endif // FORMULA

                // data or formula
                //if (!string.IsNullOrEmpty(xmlCell.data))
                //{
                //	try
                //	{
                //		SetSingleCellData(cell, xmlCell.data);
                //	}
                //	catch { }

                //	// todo: need recalc formula cells by deciding precedents
                //}

                #region Body

                // body
                if (!string.IsNullOrEmpty(xmlCell.bodyType))
                {
                    Type type;
                    if (CellTypesManager.CellTypes.TryGetValue(xmlCell.bodyType, out type))
                    {
                        try
                        {
                            cell.Body = Activator.CreateInstance(type) as ICellBody;
                        }
                        catch (Exception ex)
                        {
                            throw new ReoGridLoadException(
                                "Cannot create cell body instance from type: " + xmlCell.bodyType, ex);
                        }

                        if (type == typeof(ImageCell))
                        {
                            var leftIndex = xmlCell.data.IndexOf(',');
                            if (leftIndex > 0)
                            {
                                var mimetype = xmlCell.data.Substring(0, leftIndex);
                                if (mimetype == "image/png")
                                {
                                    var imgcode = xmlCell.data.Substring(leftIndex + 1);

                                    using (var ms = new MemoryStream(Convert.FromBase64String(imgcode)))
                                    {
#if WPF
                                        var img = new BitmapImage();
                                        img.BeginInit();
                                        img.StreamSource = ms;
                                        img.CacheOption = BitmapCacheOption.OnLoad;
                                        img.EndInit();
                                        ((ImageCell)cell.body).Image = img;
#else // WINFORM
										var img = System.Drawing.Image.FromStream(ms);
										((CellTypes.ImageCell)cell.body).Image = img;
#endif // WINFORM | WPF
                                    }

                                    cellValue = null;
                                }
                            }
                        }
                    }
                    else
                    {
                        RGFCustomBodyHandler handler;
                        if (RGFPersistenceProvider.CustomBodyTypeHandlers.TryGetValue(xmlCell.bodyType, out handler))
                            cellValue = Convert.ToString(handler.LoadData(cell, xmlCell.data));
                    }
                }

                #endregion // Body

                if (!string.IsNullOrEmpty(cellValue)) SetSingleCellData(cell, cellValue);

#if FORMULA
				if (!string.IsNullOrEmpty(formula))
				{
					try
					{
						SetCellFormula(cell, formula);
					}
					catch { }
				}
#endif // FORMULA

                // readonly flag to disable cell editing
                cell.IsReadOnly = TextFormatHelper.IsSwitchOn(xmlCell.@readonly); //Edited by Rick

                // formula trace precedents
                if (TextFormatHelper.IsSwitchOn(xmlCell.tracePrecedents))
                {
                    if (cellsTracePrecedents == null) cellsTracePrecedents = new List<Cell>();
                    cellsTracePrecedents.Add(cell);
                }

                // formula trace dependents
                if (TextFormatHelper.IsSwitchOn(xmlCell.traceDependents))
                {
                    if (cellsTraceDependents == null) cellsTraceDependents = new List<Cell>();
                    cellsTraceDependents.Add(cell);
                }
            }

            #endregion // Cells

            suspendDataChangedEvent = false;

            #region Formula

#if FORMULA
			foreach (var cell in this.formulaRanges.Keys)
			{
				// TODO: pass dirty stack?
				RecalcCell(cell);
			}

			if (cellsTracePrecedents != null)
			{
				foreach (var cell in cellsTracePrecedents)
				{
					cell.TraceFormulaPrecedents = true;
				}
			}

			if (cellsTraceDependents != null)
			{
				foreach (var cell in cellsTraceDependents)
				{
					cell.TraceFormulaDependents = true;
				}
			}
#endif // FORMULA

            #endregion // Formula

            #region Freeze

            int freezeRow = 0, freezeCol = 0;

            if (!string.IsNullOrEmpty(xmlSheet.head.freezeRow)) int.TryParse(xmlSheet.head.freezeRow, out freezeRow);
            if (!string.IsNullOrEmpty(xmlSheet.head.freezeCol)) int.TryParse(xmlSheet.head.freezeCol, out freezeCol);

            var farea = XmlFileFormatHelper.DecodeFreezeArea(xmlSheet.head.freezeArea);

            if (freezeRow > 0 || freezeCol > 0)
            {
                FreezeToCell(freezeRow, freezeCol, farea);
            }
            else
            {
                viewDirty = true;
                UpdateViewportController();
            }

            #endregion // Freeze

            #region Apply Settings

            // apply control settings
            this.settings = controlSettings;

            SettingsChanged?.Invoke(this, new SettingsChangedEventArgs
            {
                AddedSettings = this.settings
            });

            #endregion // Apply Settings

            #region Print Settings

#if PRINT
			if ((this.userPageBreakCols != null && this.userPageBreakCols.Count > 0)
				|| (this.userPageBreakRows != null && this.userPageBreakRows.Count > 0))
			{
				this.AutoSplitPage();
			}
#endif // PRINT

            #endregion // Apply Print Settings

            #region Script

#if EX_SCRIPT
			// load scripts
			// TODO: include others scripts as resource document
			if (xmlSheet.head != null && xmlSheet.head.script != null
				&& !string.IsNullOrEmpty(xmlSheet.head.script.content))
			{
				if (this.workbook == null)
				{
					throw new InvalidOperationException("Current worksheet is not attached to any workbook, loading script requires that workbook is attached. Try add worksheet into workbook firstly.");
				}

				// initialize SRM
				this.workbook.InitSRM();

				this.workbook.Script = xmlSheet.head.script.content;
			}
			else
			{
				if (this.workbook != null)
				{
					this.workbook.Script = null;
				}
			}

			if (this.Srm != null)
			{
				this.RaiseScriptEvent("onload");
			}
#endif // EX_SCRIPT

            #endregion // Script

#if DEBUG
            stop.Stop();
            Debug.WriteLine("rgf loaded: " + stop.ElapsedMilliseconds + " ms.");
#endif // DEBUG
        }

        /// <summary>
        ///     Event raised when grid loaded from file.
        /// </summary>
        public event EventHandler<FileLoadedEventArgs> FileLoaded;

        #endregion // Load

        #region Save

        /// <summary>
        ///     Save current worksheet into file.
        /// </summary>
        /// <param name="path">File path to save worksheet.</param>
        /// <returns>True if saving is successful; Otherwise return false.</returns>
        public bool Save(string path)
        {
            return Save(path, FileFormat._Auto);
        }

        /// <summary>
        ///     Save current worksheet into file.
        /// </summary>
        /// <param name="path">File path to save worksheet.</param>
        /// <param name="format">File format used to save worksheet. (Default is _Auto)</param>
        /// <returns>True if saving is successful; Otherwise return false.</returns>
        public bool Save(string path, FileFormat format)
        {
            if (format == FileFormat._Auto)
            {
                if (path.EndsWith(".xlsx", StringComparison.CurrentCultureIgnoreCase)
                    || path.EndsWith(".xlsm", StringComparison.CurrentCultureIgnoreCase))
                    format = FileFormat.Excel2007;
                else if (path.EndsWith(".xls", StringComparison.CurrentCultureIgnoreCase))
                    throw new NotSupportedException("Saving as Excel 2003 format is not supported.");
                else if (path.EndsWith(".rgf", StringComparison.CurrentCultureIgnoreCase))
                    format = FileFormat.ReoGridFormat;
                else if (path.EndsWith(".csv", StringComparison.CurrentCultureIgnoreCase)) format = FileFormat.CSV;
            }

            switch (format)
            {
                case FileFormat.ReoGridFormat:
                    return SaveRGF(path);

                case FileFormat.CSV:
                    ExportAsCSV(path);
                    return true;

                case FileFormat.Excel2007:
                    throw new NotSupportedException(
                        "Saving single worksheet is not supported, try use Workbook.Save instead.");

                default:
                    throw new NotSupportedException("Cannot determine the saving format, try explicitly specify one.");
            }
        }

        /// <summary>
        ///     Save current worksheet into file.
        /// </summary>
        /// <param name="stream">Stream to output worksheet.</param>
        /// <param name="format">File format used to save worksheet.</param>
        /// <returns>True if saving is successful; otherwise return false.</returns>
        public bool Save(Stream stream, FileFormat format = FileFormat.ReoGridFormat)
        {
            switch (format)
            {
                case FileFormat.IGXL:
                    ExportAsTxt(stream);
                    return true;

                case FileFormat.ReoGridFormat:
                    return SaveRGF(stream);

                case FileFormat.CSV:
                    ExportAsCSV(stream);
                    return true;

                case FileFormat.Excel2007:
                    throw new NotImplementedException(
                        "Saving single worksheet as Excel format is not support, try use Save method from control instance.");

                default:
                    throw new NotSupportedException();
            }
        }

        /// <summary>
        ///     Save worksheet into specified file.
        /// </summary>
        /// <param name="path">Path of file to save worksheet.</param>
        /// <returns>True if grid saved successfully.</returns>
        public bool SaveRGF(string path)
        {
            using (var fs = new FileStream(path, FileMode.Create))
            {
                var rs = SaveRGF(fs);

                // raise file saving event
                if (rs && FileSaved != null) FileSaved(this, new FileSavedEventArgs(path));

                return rs;
            }
        }

        /// <summary>
        ///     Save worksheet as RGF format into specified output stream.
        /// </summary>
        /// <param name="s">Stream to save current worksheet.</param>
        /// <returns>True if worksheet is saved successfully.</returns>
        /// <remarks>
        ///     Exceptions thrown if any errors happen during saving.
        /// </remarks>
        public bool SaveRGF(Stream s)
        {
            var editorProgram = "ReoGrid Core";

            var cas = GetType().Assembly.GetCustomAttributes(typeof(AssemblyVersionAttribute), false);

            if (cas != null && cas.Length > 0) editorProgram += " " + ((AssemblyVersionAttribute)cas[0]).Version;

            var assembly = GetType().Assembly;
            var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

            #region Head

            var body = new RGXmlSheet
            {
                head = new RGXmlHead
                {
                    cols = cols.Count,
                    rows = rows.Count,

                    defaultColumnWidth = DefaultColumnWidth,
                    defaultRowHeight = defaultRowHeight,

                    // save row-header-panel width if currently is not in auto-adjust mode
                    rowHeaderWidth = _userRowHeaderWidth ? rowHeaderWidth.ToString() : null,

                    selectionMode = selectionMode == WorksheetSelectionMode.Range
                        ? null
                        : XmlFileFormatHelper.EncodeSelectionMode(selectionMode),

                    selectionStyle = selectionStyle == WorksheetSelectionStyle.Default
                        ? null
                        : XmlFileFormatHelper.EncodeSelectionStyle(selectionStyle),

                    focusForwardDirection = selectionForwardDirection == SelectionForwardDirection.Right
                        ? null
                        : XmlFileFormatHelper.EncodeFocusForwardDirection(selectionForwardDirection),

                    focusCellStyle = focusPosStyle == FocusPosStyle.Default
                        ? null
                        : XmlFileFormatHelper.EncodeFocusPosStyle(focusPosStyle),

                    #region Settings

                    settings = new RGXmlWorksheetSetting
                    {
                        showGrid = TextFormatHelper.EncodeBool(HasSettings(WorksheetSettings.View_ShowGridLine),
                            WorksheetSettings.Default.Has(WorksheetSettings.View_ShowGridLine)),

                        showPageBreakes = TextFormatHelper.EncodeBool(
                            HasSettings(WorksheetSettings.View_ShowPageBreaks),
                            WorksheetSettings.Default.Has(WorksheetSettings.View_ShowPageBreaks)),

                        showRowHeader = TextFormatHelper.EncodeBool(HasSettings(WorksheetSettings.View_ShowRowHeader),
                            WorksheetSettings.Default.Has(WorksheetSettings.View_ShowRowHeader)),

                        showColHeader = TextFormatHelper.EncodeBool(
                            HasSettings(WorksheetSettings.View_ShowColumnHeader),
                            WorksheetSettings.Default.Has(WorksheetSettings.View_ShowColumnHeader)),

                        @readonly = TextFormatHelper.EncodeBool(HasSettings(WorksheetSettings.Edit_Readonly),
                            WorksheetSettings.Default.Has(WorksheetSettings.Edit_Readonly)),

                        allowAdjustRowHeight = TextFormatHelper.EncodeBool(
                            HasSettings(WorksheetSettings.Edit_AllowAdjustRowHeight),
                            WorksheetSettings.Default.Has(WorksheetSettings.Edit_AllowAdjustRowHeight)),

                        allowAdjustColumnWidth = TextFormatHelper.EncodeBool(
                            HasSettings(WorksheetSettings.Edit_AllowAdjustColumnWidth),
                            WorksheetSettings.Default.Has(WorksheetSettings.Edit_AllowAdjustColumnWidth)),

                        metaValue = ((long)settings).ToString()
                    },

                    #endregion // Settings

                    meta = new RGXmlMeta
                    {
                        culture = Thread.CurrentThread.CurrentCulture.Name,
                        editor = editorProgram,
                        controlVersion = fvi.FileVersion
                    },

                    script = new RGXmlScript { content = workbook.Script }
                },

                style = StyleUtility.ConvertToXmlStyle(RootStyle)
            };

            var head = body.head;

            #endregion // Head

            #region Print

#if PRINT
			if ((this.printSettings != null)
				&& (this.pageBreakRows != null && this.pageBreakCols != null
				&& this.pageBreakRows.Count > 1 && this.pageBreakCols.Count > 1))
			{
				var ps = head.printSettings = new RGXmlPrintSetting();

				// save page breaks
				if (this.userPageBreakRows != null && this.userPageBreakRows.Count > 0)
				{
					ps.pageBreakRows = TextFormatHelper.EncodeIntArray(this.userPageBreakRows);
				}
				if (this.userPageBreakCols != null && this.userPageBreakCols.Count > 0)
				{
					ps.pageBreakCols = TextFormatHelper.EncodeIntArray(this.userPageBreakCols);
				}

				if (this.printSettings != null)
				{
					ps.paperName = this.printSettings.PaperName;
					ps.pageOrder = XmlFileFormatHelper.EncodePageOrder(this.printSettings.PageOrder);
					ps.scaling = this.printSettings.PageScaling.ToString();
					ps.landscape = TextFormatHelper.EncodeBool(this.printSettings.Landscape);
					ps.paperWidth = this.printSettings.PaperWidth.ToString();
					ps.paperHeight = this.printSettings.PaperHeight.ToString();

					var margins = this.printSettings.Margins;

					ps.margins = new RGXmlMargins
					{
						left = margins.Left,
						right = margins.Right,
						top = margins.Top,
						bottom = margins.Bottom,
					};
				}
			}
#endif // PRINT

            #endregion // Print

            #region Freeze

            var freezePos = FreezePos;

            if (freezePos.Row > 0 || freezePos.Col > 0)
            {
                head.freezeRow = freezePos.Row.ToString();
                head.freezeCol = freezePos.Col.ToString();
                head.freezeArea = XmlFileFormatHelper.EncodeFreezeArea(FreezeArea);
            }

            #endregion // Freeze

            #region Outlines

#if OUTLINE
			// outlines
			if (this.outlines != null)
			{
				Action<OutlineCollection<ReoGridOutline>, List<RGXmlOutline>> addOutliens = (outlines, xmlOutlines) =>
				{
					outlines.IterateOutlines((outline) =>
					{
						xmlOutlines.Add(new RGXmlOutline
						{
							start = outline.Start,
							count = outline.Count,
							collapsed = outline.InternalCollapsed,
						});

						return true;
					});
				};

				var rowOutlines = GetOutlines(RowOrColumn.Row);
				if (rowOutlines != null)
				{
					if (head.outlines == null)
					{
						head.outlines = new RGXmlOutlineList();
					}

					head.outlines.rowOutlines = new List<RGXmlOutline>();
					addOutliens(rowOutlines, head.outlines.rowOutlines);
				}

				var colOutlines = GetOutlines(RowOrColumn.Column);
				if (colOutlines != null)
				{
					if (head.outlines == null)
					{
						head.outlines = new RGXmlOutlineList();
					}

					head.outlines.colOutlines = new List<RGXmlOutline>();
					addOutliens(colOutlines, head.outlines.colOutlines);
				}
			}
#endif // OUTLINE

            #endregion // Outlines

            #region Named Ranges

            // named ranges
            if (registeredNamedRanges.Count > 0)
            {
                head.namedRanges = new List<RGXmlNamedRange>();

                head.namedRanges.AddRange(
                    from nr in registeredNamedRanges.Values
                    select new RGXmlNamedRange
                    {
                        name = nr.Name,
                        comment = nr.Comment,
                        address = nr.Position.ToAddress()
                    });
            }

            #endregion // Named Ranges

            var maxRow = MaxContentRow;
            var maxCol = MaxContentCol;

            #region Rows

            // row-headers
            foreach (var r in rows)
            {
                var toSave = r.InnerHeight != defaultRowHeight || r.InnerStyle != null
                                                               || !r.IsAutoHeight || r.LastHeight != 0
                                                               || !string.IsNullOrEmpty(r.Text) || r.TextColor != null;

                WorksheetRangeStyle checkedStyle = null;

                checkedStyle = StyleUtility.DistinctStyle(r.InnerStyle, RootStyle);

                if (toSave || checkedStyle != null)
                    body.rows.Add(new RGXmlRowHead
                    {
                        row = r.Row,
                        height = r.InnerHeight,

                        lastHeight = r.LastHeight == 0 ? null : r.LastHeight.ToString(),
                        autoHeight = r.IsAutoHeight ? null : TextFormatHelper.EncodeBool(r.IsAutoHeight),

                        text = string.IsNullOrEmpty(r.Text) ? null : r.Text,
                        textColor = r.TextColor == null ? null : TextFormatHelper.EncodeColor(r.TextColor.Value),

                        style = StyleUtility.ConvertToXmlStyle(checkedStyle)
                    });
            }

            #endregion // Rows

            #region Columns

            // col-headers
            foreach (var c in cols)
            {
                var toSave = c.InnerWidth != DefaultColumnWidth || c.InnerStyle != null
                                                                || !c.IsAutoWidth || c.LastWidth != 0
                                                                || !string.IsNullOrEmpty(c.Text) || c.TextColor != null
                                                                || c.DefaultCellBody != null;

                WorksheetRangeStyle checkedStyle = null;

                checkedStyle = StyleUtility.DistinctStyle(c.InnerStyle, RootStyle);

                if (toSave || checkedStyle != null)
                    body.cols.Add(new RGXmlColHead
                    {
                        col = c.Col,
                        width = c.InnerWidth,

                        lastWidth = c.LastWidth == 0 ? null : c.LastWidth.ToString(),
                        autoWidth = c.IsAutoWidth ? null : TextFormatHelper.EncodeBool(c.IsAutoWidth),

                        text = string.IsNullOrEmpty(c.Text) ? null : c.Text,
                        textColor = c.TextColor == null ? null : TextFormatHelper.EncodeColor(c.TextColor.Value),

                        style = StyleUtility.ConvertToXmlStyle(checkedStyle),

                        defaultCellBody = c.DefaultCellBody == null ? null : c.DefaultCellBody.FullName
                    });
            }

            #endregion // Columns

            #region Borders

            // h-borders
            for (var r = 0; r <= maxRow + 1; r++)
            for (var c = 0; c <= maxCol;)
            {
                var cellBorder = hBorders[r, c];

                if (cellBorder != null && cellBorder.Span > 0 && cellBorder.Style != null
                    && cellBorder.Style.Style != BorderLineStyle.None)
                {
                    body.hborder.Add(new RGXmlHBorder(r, c, cellBorder.Span, cellBorder.Style, cellBorder.Pos));
                    c += cellBorder.Span;
                }
                else
                {
                    c++;
                }
            }

            // v-borders
            for (var c = 0; c <= maxCol + 1; c++)
            for (var r = 0; r <= maxRow;)
            {
                var cellBorder = vBorders[r, c];

                if (cellBorder != null && cellBorder.Span > 0 && cellBorder.Style != null
                    && cellBorder.Style.Style != BorderLineStyle.None)
                {
                    body.vborder.Add(new RGXmlVBorder(r, c, cellBorder.Span, cellBorder.Style, cellBorder.Pos));
                    r += cellBorder.Span;
                }
                else
                {
                    r++;
                }
            }

            #endregion // Borders

            #region Cells

            // cells
            cells.Iterate(0, 0, maxRow + 1, maxCol + 1, true, (r, c, cell) =>
            {
                if (cell.IsValidCell || cell.IsStartMergedCell)
                {
                    var addCell = false;

                    if (cell.InnerData != null || cell.Rowspan > 1 || cell.Colspan > 1 || cell.body != null)
                        addCell = true;

                    RGXmlCellStyle xmlStyle = null;

                    if (cell.StyleParentKind == StyleParentKind.Own)
                    {
                        xmlStyle = StyleUtility.ConvertToXmlStyle(
                            StyleUtility.CheckAndRemoveCellStyle(this, cell));

                        if (xmlStyle != null) addCell = true;
                    }

                    #region Data Format

                    RGXmlCellDataFormatArgs xmlFormatArgs = null;
                    if (cell.DataFormat != CellDataFormatFlag.General)
                    {
                        addCell = true;

                        switch (cell.DataFormat)
                        {
                            case CellDataFormatFlag.Number:
                                var nargs = (NumberDataFormatter.NumberFormatArgs)cell.DataFormatArgs;
                                xmlFormatArgs = new RGXmlCellDataFormatArgs();
                                xmlFormatArgs.decimalPlaces = nargs.DecimalPlaces.ToString();
                                xmlFormatArgs.negativeStyle =
                                    XmlFileFormatHelper.EncodeNegativeNumberStyle(nargs.NegativeStyle);
                                xmlFormatArgs.useSeparator = TextFormatHelper.EncodeBool(nargs.UseSeparator);
                                break;

                            case CellDataFormatFlag.DateTime:
                                var dargs = (DateTimeDataFormatter.DateTimeFormatArgs)cell.DataFormatArgs;
                                xmlFormatArgs = new RGXmlCellDataFormatArgs();
                                xmlFormatArgs.culture = dargs.CultureName;
                                xmlFormatArgs.pattern = dargs.Format;
                                break;

                            case CellDataFormatFlag.Currency:
                                var cargs = (CurrencyDataFormatter.CurrencyFormatArgs)cell.DataFormatArgs;
                                xmlFormatArgs = new RGXmlCellDataFormatArgs();
                                xmlFormatArgs.decimalPlaces = cargs.DecimalPlaces.ToString();
                                xmlFormatArgs.culture = cargs.CultureEnglishName;
                                xmlFormatArgs.negativeStyle =
                                    XmlFileFormatHelper.EncodeNegativeNumberStyle(cargs.NegativeStyle);
                                xmlFormatArgs.pattern = cargs.PrefixSymbol + "," + cargs.PostfixSymbol;
                                break;

                            case CellDataFormatFlag.Percent:
                                var pargs = (NumberDataFormatter.NumberFormatArgs)cell.DataFormatArgs;
                                xmlFormatArgs = new RGXmlCellDataFormatArgs();
                                xmlFormatArgs.decimalPlaces = pargs.DecimalPlaces.ToString();
                                break;
                        }
                    }

                    #endregion // Data Format

                    #region Cell

                    if (addCell)
                    {
                        var xmlCell = new RGXmlCell
                        {
                            row = r,
                            col = c,
                            colspan = cell.Colspan == 1 ? null : cell.Colspan.ToString(),
                            rowspan = cell.Rowspan == 1 ? null : cell.Rowspan.ToString(),
                            style = xmlStyle,
                            dataFormat = XmlFileFormatHelper.EncodeCellDataFormat(cell.DataFormat),
                            dataFormatArgs = xmlFormatArgs == null || xmlFormatArgs.IsEmpty ? null : xmlFormatArgs,
                            @readonly = cell.IsReadOnly == false
                                ? null
                                : TextFormatHelper.EncodeBool(cell.IsReadOnly), // Edited by Rick

                            bodyType = cell.body == null ? null : cell.body.GetType().Name,

#if FORMULA
							tracePrecedents =
 cell.TraceFormulaPrecedents ? TextFormatHelper.EncodeBool(cell.TraceFormulaPrecedents) : null,
							traceDependents =
 cell.TraceFormulaDependents ? TextFormatHelper.EncodeBool(cell.TraceFormulaDependents) : null,
#endif // FORMULA
                        };

                        if (cell.HasFormula || !(cell.InnerData is bool))
                        {
                            xmlCell.data = Convert.ToString(cell.InnerData);
                            xmlCell.formula = cell.HasFormula ? new RGXmlCellFormual { val = cell.InnerFormula } : null;
                        }
                        else if (cell.InnerData is bool)
                        {
                            xmlCell.formula = new RGXmlCellFormual { val = (bool)cell.InnerData ? "True" : "False" };
                        }

                        if (cell.body != null)
                        {
                            if (cell.body is ImageCell)
                            {
                                var imageBody = (ImageCell)cell.body;
                                using (var ms = new MemoryStream(4096))
                                {
#if WPF
                                    var img = ((ImageCell)cell.body).Image;
                                    var enc = new PngBitmapEncoder();
                                    try
                                    {
                                        enc.Frames.Add(BitmapFrame.Create((BitmapSource)img));
                                    }
                                    catch (NotSupportedException)
                                    {
                                    }

                                    enc.Save(ms);
#else // WINFORM
									imageBody.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

#endif // WINFORM | WPF
                                    xmlCell.data = "image/png," + Convert.ToBase64String(ms.ToArray());
                                }
                            }
                            else
                            {
                                string typeIdentifier;
                                if (RGFPersistenceProvider.CustomBodyTypeIdentifiers.TryGetValue(cell.body.GetType(),
                                        out typeIdentifier))
                                {
                                    RGFCustomBodyHandler handler;
                                    if (RGFPersistenceProvider.CustomBodyTypeHandlers.TryGetValue(typeIdentifier,
                                            out handler))
                                    {
                                        xmlCell.bodyType = typeIdentifier;
                                        xmlCell.data = handler.SaveData(cell);
                                    }
                                }
                            }
                        }

                        body.cells.Add(xmlCell);
                    }

                    #endregion // Cell
                }

                return 1;
            });

            #endregion // Cells

            var xmlWriter = new XmlSerializer(typeof(RGXmlSheet));
            xmlWriter.Serialize(s, body);

            return true;
        }

        /// <summary>
        ///     Event raised when worksheet saved into a file.
        /// </summary>
        public event EventHandler<FileSavedEventArgs> FileSaved;

        #endregion // Save
    }
}