﻿#define WPF

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Controls;
using SpreedSheet.CellTypes;
using SpreedSheet.Core;
using SpreedSheet.Core.Workbook;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.Common;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WINFORM

namespace unvell.ReoGrid.Data
{
    #region Auot Filter

    /// <summary>
    ///     Built-in auto column filter
    /// </summary>
    public class AutoColumnFilter
    {
        private FilterColumnCollection columnCollection;

        private AutoColumnFilterUI columnFilterUIFlag;

        //internal RangePosition ApplyRange
        //{
        //	get
        //	{
        //		return new RangePosition(this.titleRows, this.startCol, -1, this.endCol - this.startCol + 1);
        //	}
        //}

        internal AutoColumnFilter(Worksheet worksheet, /*int startCol, int endCol, int titleRows = 1*/
            RangePosition range)
        {
//#if DEBUG
//			Debug.Assert(endCol >= startCol);
//#endif

//			if (startCol < 0) throw new ArgumentOutOfRangeException("start number of column out of spreadsheet");
//			if (startCol >= worksheet.ColumnCount) throw new ArgumentOutOfRangeException("start number of column out of spreadsheet");

//			if (endCol < 0) throw new ArgumentOutOfRangeException("end number of column out of spreadsheet");
//			if (endCol >= worksheet.ColumnCount) throw new ArgumentOutOfRangeException("end number of column out of spreadsheet");

//			if (endCol < startCol) throw new ArgumentOutOfRangeException("end number of column must be greater than start number of column");

            Worksheet = worksheet;
            ApplyRange = range;
            //this.startCol = startCol;
            //this.endCol = endCol;
            //this.titleRows = titleRows;
        }
        //internal Worksheet innerWorksheet { get; set; }

        /// <summary>
        ///     Get instance of current attached worksheet.
        /// </summary>
        public Worksheet Worksheet { get; private set; }

        public RangePosition ApplyRange { get; }

        /// <summary>
        ///     Get the columns of this filter.
        /// </summary>
        public FilterColumnCollection Columns
        {
            get
            {
                if (columnCollection == null) columnCollection = new FilterColumnCollection(this);
                return columnCollection;
            }
        }

        /// <summary>
        ///     Apply filter to update worksheet.
        /// </summary>
        public void Apply()
        {
            if (Worksheet == null) return;

            //int endRow = this.Worksheet.MaxContentRow;

            Worksheet.DoFilter(ApplyRange, r =>
            {
                for (var c = ApplyRange.Col; c <= ApplyRange.EndCol;)
                {
                    var columnHeader = Worksheet.RetrieveColumnHeader(c);
                    if (columnHeader == null)
                    {
                        c++;
                        continue;
                    }

                    var columnFilterBody = columnHeader.Body as AutoColumnFilterBody;
                    if (columnFilterBody == null || columnFilterBody.IsSelectAll)
                    {
                        c++;
                        continue;
                    }

                    var cell = Worksheet.GetCell(r, c);
                    if (cell == null)
                    {
                        c++;
                        continue;
                    }

                    var text = cell.DisplayText;
                    if (string.IsNullOrEmpty(text)) text = LanguageResource.Filter_Blanks;

                    if (!columnFilterBody.SelectedTextItems.Contains(text))
                        // hide the row
                        return false;

                    c += cell.Colspan;
                }

                // show the row
                return true;
            });
        }

        /// <summary>
        ///     Event raised when column filter button has been clicked
        /// </summary>
        public event EventHandler FilterButtonPressed;

        internal bool RaiseFilterButtonPress(AutoColumnFilterBody headerBody, Point point)
        {
            if (headerBody.ColumnHeader == null) return false;

            if (FilterButtonPressed != null)
            {
                var arg = new FilterButtonPressedEventArgs(headerBody.ColumnHeader);

                FilterButtonPressed(this, arg);

                if (arg.IsCancelled) return true;
            }

            if (columnFilterUIFlag == AutoColumnFilterUI.DropdownButtonAndPanel)
            {
#if WINFORM
				unvell.ReoGrid.WinForm.ColumnFilterContextMenu.ShowFilterPanel(headerBody, (System.Drawing.Point)point);
#elif WPF
                var ctx = new ContextMenu();
                ctx.Items.Add(new MenuItem { Header = "Item" });
                ctx.IsOpen = true;
#endif // WPF
                return true;
            }

            return false;
        }

        private void CreateFilterHeaders(int start, int end)
        {
            try
            {
                Worksheet.ControlAdapter.ChangeCursor(CursorStyle.Busy);

                for (var i = start; i <= end; i++)
                {
                    var header = Worksheet.RetrieveColumnHeader(i);
                    header.Body = new AutoColumnFilterBody(this, header);
                }
            }
            finally
            {
                Worksheet.ControlAdapter.ChangeCursor(CursorStyle.PlatformDefault);
            }
        }

        private void RemoveFilterHeader(int start, int end)
        {
            if (Worksheet.cols.Count <= 0) return;

            for (var i = start; i <= end; i++)
            {
                var header = Worksheet.RetrieveColumnHeader(i);
                var body = header.Body as AutoColumnFilterBody;

                if (body != null)
                {
#if WINFORM
					if (body.ContextMenuStrip != null)
					{
						body.ContextMenuStrip.Dispose();
					}
#elif WPF
                    // todo
#endif

                    header.Body = null;
                }
            }
        }

        /// <summary>
        ///     Attach filter to specified worksheet
        /// </summary>
        /// <param name="worksheet">instance of worksheet to be attached</param>
        /// <param name="uiFlag">Flags to decide which styles of GUI to be dispalyed (default is DropdownButtonAndPanel style)</param>
        public void Attach(Worksheet worksheet, AutoColumnFilterUI uiFlag = AutoColumnFilterUI.DropdownButtonAndPanel)
        {
            if (worksheet == null) throw new ArgumentNullException("cannot attach to null worksheet", "worksheet");

            Worksheet = worksheet;
            columnFilterUIFlag = uiFlag;

            CreateFilterHeaders(ApplyRange.Col, ApplyRange.EndCol);

            worksheet.ColumnsInserted += worksheet_ColumnsInserted;
        }

        /// <summary>
        ///     Detach filter from specified worksheet
        /// </summary>
        public void Detach()
        {
            if (Worksheet != null)
            {
                Worksheet.ColumnsInserted -= worksheet_ColumnsInserted;

                Worksheet.ShowRows(ApplyRange.Row, ApplyRange.Rows);

                RemoveFilterHeader(ApplyRange.Col, ApplyRange.EndCol);

                Worksheet.RequestInvalidate();
                Worksheet = null;
            }
        }

        private void worksheet_ColumnsInserted(object sender, ColumnsInsertedEventArgs e)
        {
            if (e.Index < ApplyRange.Col)
            {
                //RemoveFilterHeader(this.startCol, e.Index - 1);
                ApplyRange.Offset(0, e.Count);
            }
            else if (e.Index > ApplyRange.Col && e.Index <= ApplyRange.EndCol)
            {
                CreateFilterHeaders(e.Index, e.Index + e.Count - 1);
                ApplyRange.SetCols(ApplyRange.Cols + e.Count);
            }
        }

        #region FilterColumnCollection

        /// <summary>
        ///     Collection of column filters
        /// </summary>
        public class FilterColumnCollection : IEnumerable<AutoColumnFilterBody>
        {
            private readonly AutoColumnFilter filter;

            internal FilterColumnCollection(AutoColumnFilter filter)
            {
                this.filter = filter;
            }

            /// <summary>
            ///     Get the column filter by specified index.
            /// </summary>
            /// <param name="index">number of column to get column filter.</param>
            /// <returns>instance of column filter, which contains the candidates list and selected items by user.</returns>
            public AutoColumnFilterBody this[int index]
            {
                get
                {
                    if (index < filter.ApplyRange.Col || index > filter.ApplyRange.EndCol
                                                      || index < 0 || index >= filter.Worksheet.ColumnCount)
                        throw new ArgumentOutOfRangeException("index",
                            "Number of column to find the filter out of the valid range");

                    return filter.Worksheet.RetrieveColumnHeader(index).Body as AutoColumnFilterBody;
                }
            }

            /// <summary>
            ///     Get the column filter by specified address code of column (A, TZ, etc.)
            /// </summary>
            /// <param name="columnCode">the alphabet of address code used to locate a column in spreadsheet</param>
            /// <returns>instance of column filter, which contains the candidates list and selected items by user</returns>
            public AutoColumnFilterBody this[string columnCode]
            {
                get
                {
                    var index = RGUtility.GetNumberOfChar(columnCode);
                    return this[index];
                }
            }

            /// <summary>
            ///     Get all column filter header body from this filter.
            /// </summary>
            /// <returns></returns>
            public IEnumerator<AutoColumnFilterBody> GetEnumerator()
            {
                return GetEnum();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnum();
            }

            private IEnumerator<AutoColumnFilterBody> GetEnum()
            {
                for (var i = filter.ApplyRange.Col; i <= filter.ApplyRange.EndCol; i++)
                {
                    var header = filter.Worksheet.RetrieveColumnHeader(i);
                    if (header == null) continue;
                    var body = header.Body as AutoColumnFilterBody;
                    if (body == null) continue;
                    yield return body;
                }
            }
        }

        #endregion // FilterColumnCollection

        #region AutoColumnFilterBody

        public class AutoColumnFilterBody : IHeaderBody
        {
            internal AutoColumnFilter autoFilter;

            /// <summary>
            ///     Column header instance
            /// </summary>
            public ColumnHeader ColumnHeader { get; }

            internal AutoColumnFilterBody(AutoColumnFilter autoFilter, ColumnHeader header)
            {
                this.autoFilter = autoFilter;
                ColumnHeader = header;
                DataDirty = true;
                IsSelectAll = true;
            }

            internal bool IsDropdown { get; set; }

            /// <summary>
            ///     Repaint filter header body
            /// </summary>
            /// <param name="dc">ReoGrid drawing context</param>
            /// <param name="headerSize">Header size</param>
            public void OnPaint(CellDrawingContext dc, Size headerSize)
            {
                var controlStyle = dc.Worksheet.controlAdapter.ControlStyle;

                if (autoFilter.columnFilterUIFlag == AutoColumnFilterUI.DropdownButton
                    || autoFilter.columnFilterUIFlag == AutoColumnFilterUI.DropdownButtonAndPanel)
                {
                    var bounds = GetColumnFilterButtonRect(headerSize);

                    var color1 = controlStyle.GetColHeadStartColor(false, isInvalid: false,
                        isSelected: IsDropdown, isFullSelected: false);

                    var color2 = controlStyle.GetColHeadEndColor(false, isInvalid: false,
                        isSelected: IsDropdown, isFullSelected: false);

                    var g = dc.Graphics;

                    g.FillRectangleLinear(color1, color2, 90f, bounds);

                    g.DrawRectangle(bounds, StaticResources.SystemColor_ControlDark);

                    GraphicsToolkit.FillTriangle(dc.Graphics.PlatformGraphics,
                        Math.Min(7 * dc.Worksheet.renderScaleFactor, 7.0f),
                        new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2));
                }
            }

            /// <summary>
            ///     Handling mouse-down process
            /// </summary>
            /// <param name="headerSize">Header size</param>
            /// <param name="e">Argument of mouse-down event</param>
            /// <returns>True if event has been handled; otherwise return false</returns>
            public bool OnMouseDown(Size headerSize, WorksheetMouseEventArgs e)
            {
                if (autoFilter == null
                    || autoFilter.columnFilterUIFlag == AutoColumnFilterUI.NoGUI
                    || ColumnHeader == null || ColumnHeader.Worksheet == null) return false;

                if (IsMouseInButton(headerSize, e.RelativePosition))
                    return autoFilter.RaiseFilterButtonPress(this, e.AbsolutePosition);
                return false;
            }

            /// <summary>
            ///     Handling mouse-move process
            /// </summary>
            /// <param name="headerSize">Header size</param>
            /// <param name="e">Argument of mouse-move event</param>
            /// <returns>True if event has been handled, otherwise return false</returns>
            public bool OnMouseMove(Size headerSize, WorksheetMouseEventArgs e)
            {
                if (autoFilter == null
                    || autoFilter.columnFilterUIFlag == AutoColumnFilterUI.NoGUI
                    || ColumnHeader == null || ColumnHeader.Worksheet == null) return false;

                if (IsMouseInButton(headerSize, e.RelativePosition))
                {
                    e.CursorStyle = CursorStyle.Hand;
                    return true;
                }

                return false;
            }

            private bool IsMouseInButton(Size headerSize, Point position)
            {
                var bounds = GetColumnFilterButtonRect(headerSize);
                //Console.WriteLine(bounds.ToString() + " ---- " + position.ToString());
                return bounds.Contains(position);
            }

            internal Rectangle GetColumnFilterButtonRect(Size size)
            {
                var sheet = ColumnHeader.Worksheet;

                var scale = sheet.renderScaleFactor;

                var bounds = new Rectangle(0, 0, Math.Min(Math.Min(size.Width - 2, 18f * scale), 20),
                    Math.Min(Math.Min(size.Height - 2, 18 * scale), 20));
                bounds.X = size.Width - bounds.Width - 2;
                bounds.Y = (size.Height - bounds.Height) / 2 - 1;

                return bounds;
            }

#if WINFORM
			/// <summary>
			/// Get or set the context menu strip of column filter.
			/// </summary>
			public System.Windows.Forms.ContextMenuStrip ContextMenuStrip { get; set; }
#elif WPF

#endif

            internal List<string> selectedTextItems = new List<string>();

            /// <summary>
            ///     Get or set whether or not this column is marked as SelectAll.
            ///     The column is marked as SelectAll will be ignored during filter.
            /// </summary>
            public bool IsSelectAll { get; set; }

            private TextFilterCollection textItemsCollection;

            /// <summary>
            ///     Collection of selected items
            /// </summary>
            public TextFilterCollection SelectedTextItems
            {
                get
                {
                    if (textItemsCollection == null) textItemsCollection = new TextFilterCollection(this);

                    return textItemsCollection;
                }
            }

            #region TextFilterCollection

            /// <summary>
            ///     Collection of selected items for column filter
            /// </summary>
            public class TextFilterCollection : ICollection<string>
            {
                private readonly AutoColumnFilterBody columnFilter;

                internal TextFilterCollection(AutoColumnFilterBody column)
                {
                    columnFilter = column;
                }

                /// <summary>
                ///     Check whether or not a specified item is selected by user.
                /// </summary>
                /// <param name="item">Item to be checked.</param>
                /// <returns>True if item is selected by user; Otherwise return false;</returns>
                public bool this[string item]
                {
                    get { return columnFilter.selectedTextItems.Contains(item); }
                    set
                    {
                        if (value)
                            Add(item);
                        else
                            Remove(item);
                    }
                }

                /// <summary>
                ///     Get the enumeration of items from this filter.
                /// </summary>
                /// <returns>Enumeration of items from this filter.</returns>
                public IEnumerator<string> GetEnumerator()
                {
                    return columnFilter.selectedTextItems.GetEnumerator();
                }

                IEnumerator IEnumerable.GetEnumerator()
                {
                    return columnFilter.selectedTextItems.GetEnumerator();
                }

                /// <summary>
                ///     Add selected item
                /// </summary>
                /// <param name="item">item to be added</param>
                public void Add(string item)
                {
                    if (!columnFilter.selectedTextItems.Contains(item))
                    {
                        columnFilter.selectedTextItems.Add(item);
                        columnFilter.IsSelectAll = false;
                    }
                }

                /// <summary>
                ///     Clear all selected items
                /// </summary>
                public void Clear()
                {
                    columnFilter.selectedTextItems.Clear();
                    columnFilter.IsSelectAll = false;
                }

                /// <summary>
                ///     Check whether the specified item is contained in selected items
                /// </summary>
                /// <param name="item">item to be checked</param>
                /// <returns>true if specified item has been contained in selected items, otherwise return false</returns>
                public bool Contains(string item)
                {
                    return columnFilter.selectedTextItems.Contains(item);
                }

                /// <summary>
                ///     Copy all selected items into specified array
                /// </summary>
                /// <param name="array">array to be filled</param>
                /// <param name="arrayIndex">number of element start to copy</param>
                public void CopyTo(string[] array, int arrayIndex)
                {
                    columnFilter.selectedTextItems.CopyTo(array, arrayIndex);
                }

                /// <summary>
                ///     Get number of selected items
                /// </summary>
                public int Count
                {
                    get { return columnFilter.selectedTextItems.Count; }
                }

                /// <summary>
                ///     Check whether or not the collection of selection items is read-only
                /// </summary>
                public bool IsReadOnly
                {
                    get { return false; }
                }

                /// <summary>
                ///     Remove specified item from selected items
                /// </summary>
                /// <param name="item">item to be removed</param>
                /// <returns>true if item exist and has been removed successfully</returns>
                public bool Remove(string item)
                {
                    columnFilter.IsSelectAll = false;
                    return columnFilter.selectedTextItems.Remove(item);
                }

                /// <summary>
                ///     Add entire specified array or enumerable list into selected items
                /// </summary>
                /// <param name="items">list, array or other enumerable collection to be added</param>
                public void AddRange(IEnumerable<string> items)
                {
                    columnFilter.selectedTextItems.AddRange(items);
                    columnFilter.IsSelectAll = false;
                }
            }

            #endregion // TextFilterCollection

            /// <summary>
            ///     Get distinct items from spreadsheet on current column
            /// </summary>
            /// <returns></returns>
            public List<string> GetDistinctItems()
            {
                if (ColumnHeader == null || ColumnHeader.Worksheet == null) return null;

                var items = new List<string>();

                var maxRow = ColumnHeader.Worksheet.MaxContentRow;

                ColumnHeader.Worksheet.IterateCells(autoFilter.ApplyRange.Row,
                    ColumnHeader.Index, autoFilter.ApplyRange.Rows, 1, true,
                    (r, c, cell) =>
                    {
                        var str = cell.DisplayText;
                        if (string.IsNullOrEmpty(str)) str = LanguageResource.Filter_Blanks;

                        if (!items.Contains(str)) items.Add(str);

                        return true;
                    });

                items.Sort();

                return items;
            }

            internal bool DataDirty { get; set; }

            /// <summary>
            ///     Invoked when spreadsheet data changed on this column
            /// </summary>
            /// <param name="startRow">zero-based number of first row that data has been changed</param>
            /// <param name="endRow">zero-based number of last row that data has been changed</param>
            public void OnDataChange(int startRow, int endRow)
            {
                DataDirty = true;
            }
        }

        #endregion // AutoColumnFilterBody

        #region IColumnFilter Members

        #endregion
    }

    #endregion // Auto Filter

    /// <summary>
    ///     Flag to create UI of column filter
    /// </summary>
    public enum AutoColumnFilterUI
    {
        /// <summary>
        ///     Do not create any GUI
        /// </summary>
        NoGUI,

        /// <summary>
        ///     Only create a dropdown button on header
        /// </summary>
        DropdownButton,

        /// <summary>
        ///     Create both dropdown button and built-in panel
        /// </summary>
        DropdownButtonAndPanel
    }

    /// <summary>
    ///     Event raised when auto filter button was pressed by user
    /// </summary>
    public class FilterButtonPressedEventArgs : EventArgs
    {
        /// <summary>
        ///     Create filter button pressed event arguments with instance of column header
        /// </summary>
        /// <param name="columnHeader"></param>
        public FilterButtonPressedEventArgs(ColumnHeader columnHeader)
        {
            ColumnHeader = columnHeader;
        }

        /// <summary>
        ///     Get the instance of column header
        /// </summary>
        public ColumnHeader ColumnHeader { get; }

        /// <summary>
        ///     Set this flag to prevent open the built-in popup menu
        /// </summary>
        public bool IsCancelled { get; set; }
    }
}