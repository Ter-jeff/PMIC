#define WPF


#if DEBUG
#endif // DEBUG

#if WINFORM || ANDROID
using RGFloat = System.Single;
#elif WPF
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using SpreedSheet.CellTypes;
using SpreedSheet.Core;
using SpreedSheet.Core.Enum;
using unvell.Common;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Utility;
using RGFloat = System.Double;

#elif iOS
using RGFloat = System.Double;
#endif // WINFORM

#if EX_SCRIPT
using unvell.ReoScript;
using unvell.ReoGrid.Script;
#endif // EX_SCRIPT

#if FORMULA
using unvell.ReoGrid.Formula;
#endif // FORMULA

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        /// <summary>
        ///     Determine whether or not to suspend all data changing event. Suspend this event when
        ///     update lot of data into spreadsheet will help to speed up the data updating process.
        ///     Note: make sure this flag can be restored for event invoke.
        /// </summary>
        private bool suspendDataChangedEvent;

        /// <summary>
        ///     Suspend all data changing events, from both cell and range.
        /// </summary>
        public void SuspendDataChangedEvents()
        {
            suspendDataChangedEvent = true;
        }

        /// <summary>
        ///     Resume all data changing events.
        /// </summary>
        public void ResumeDataChangedEvents()
        {
            suspendDataChangedEvent = false;
            RequestInvalidate();
        }

        internal void RaiseCellDataChangedEvent(Cell cell)
        {
            if (CellDataChanged != null) CellDataChanged(this, new CellEventArgs(cell));

#if EX_SCRIPT
			RaiseScriptEvent("ondatachange", new RSCellObject(this, cell.InternalPos, cell));
#endif
        }

        /// <summary>
        ///     Event raised when any data has been changed
        /// </summary>
        public event EventHandler<CellEventArgs> CellDataChanged;

        #region Set Data

        /// <summary>
        ///     Set data of cell at specified position on worksheet.
        /// </summary>
        /// <param name="addressOrName">Address or name to locate the cell.</param>
        /// <param name="data">Data to be set.</param>
        public void SetCellData(string addressOrName, object data)
        {
            if (CellPosition.IsValidAddress(addressOrName))
            {
                SetCellData(new CellPosition(addressOrName), data);
            }
            else
            {
                NamedRange range;
                if (registeredNamedRanges.TryGetValue(addressOrName, out range))
                    SetCellData(range.StartPos, data);
                else
                    throw new InvalidAddressException(addressOrName);
            }
        }

        /// <summary>
        ///     Set data of cell at specified position on worksheet.
        /// </summary>
        /// <param name="pos">Position of cell to set data.</param>
        /// <param name="data">Data of cell to be set.</param>
        public void SetCellData(CellPosition pos, object data)
        {
            SetCellData(pos.Row, pos.Col, data);
        }

        /// <summary>
        ///     Set data of cell at specified position on worksheet.
        /// </summary>
        /// <param name="row">Index of row of specified cell.</param>
        /// <param name="col">Index of column of specified cell.</param>
        /// <param name="data">Data of cell.</param>
        public void SetCellData(int row, int col, object data)
        {
            if (row < 0 || row > this.rows.Count - 1)
                throw new ArgumentOutOfRangeException("row",
                    "Number of row is out of the maximum rows, use either AppendRows or Resize to expend this worksheet.");

            if (col < 0 || col > this.cols.Count - 1)
                throw new ArgumentOutOfRangeException("col",
                    "Number of column is out of maximum columns, use either AppendCols or Resize to expend this worksheet.");

            if (data is Array)
            {
                var arr = (Array)data;

                if (arr.Rank == 1)
                {
                    for (var c = col; c < Math.Min(col + arr.Length, cols.Count); c++)
                        SetCellData(row, c, arr.GetValue(c - col));
                }
                else if (arr.Rank == 2)
                {
                    var rows = arr.GetLength(0);
                    var cols = arr.GetLength(1);
                    SetRangeData(new RangePosition(row, col, rows, cols), arr);
                }
                else
                {
                    throw new ArgumentException("Array with more than 2 ranks is not supported.");
                }
            }
            else if (data is IEnumerable<object>)
            {
                var elements = (IEnumerable<object>)data;

                var c = col;
                foreach (var ele in elements)
                {
                    SetCellData(row, c, ele);
                    c++;
                    if (c >= cols.Count) break;
                }
            }
            else if (data is PartialGrid)
            {
                var subgrid = (PartialGrid)data;

                var range = new RangePosition(row, col, subgrid.Rows, subgrid.Columns);
                SetPartialGrid(range, subgrid);
            }
            else if (data is DataTable)
            {
                var dt = (DataTable)data;
                SetRangeData(new RangePosition(row, col, dt.Rows.Count, dt.Columns.Count), dt);
            }
            else
            {
                var cell = cells[row, col];

                // both data and cell is null, then no need to update
                if ((data != null || cell != null)

                    // if cell is not null, and it is valid (not merged by other cells), then need to update
                    && (cell == null || cell.IsValidCell))
                    SetSingleCellData(CreateAndGetCell(row, col), data);
            }
        }

        /// <summary>
        ///     Set data of cell at specified position on worksheet.
        /// </summary>
        /// <param name="cell">Instance of cell.</param>
        /// <param name="data">Data to be set.</param>
        internal void SetSingleCellData(Cell cell, object data)
        {
            // set cell body
            if (data is ICellBody)
            {
                SetCellBody(cell, (ICellBody)data);

                data = cell.InnerData;
            }

            if (data is string || data is StringBuilder
#if EX_SCRIPT
 || data is StringObject
#endif // EX_SCRIPT
               )
            {
                var str = data is string ? (string)data : Convert.ToString(data);

                // cell data processed as plain-text
                if (str.Length > 1)
                {
                    if (str[0] == '\'')
                    {
#if FORMULA
						// clear old references
						ClearCellReferenceList(cell);

						// clear dependents arrows
						RemoveCellTraceDependents(cell);

						// clear precedents arrow
						RemoveCellTracePrecedents(cell);

						// clear formula status
						cell.formulaStatus = FormulaStatus.Normal;

#endif // FORMULA

                        cell.InnerData = data;
                        cell.InnerDisplay = str.Substring(1);

                        AfterCellDataUpdate(cell);
                        return;
                    }

#if FORMULA
					if (str[0] == '=')
					{
						SetCellFormula(cell, str.Substring(1));

						try
						{
							RecalcCell(cell);
						}
						catch (Exception ex)
						{
							this.NotifyExceptionHappen(ex);
						}

						return;
					}
#endif // FORMULA
                }
            }

            // experimental: directly set an image as cell data
            //
            //else if (data is System.Drawing.Image)
            //{
            //	if (cell.body == null)
            //	{
            //		cell.Body = new ImageCell((System.Drawing.Image)data);
            //	}
            //	else if (cell.body is ImageCell)
            //	{
            //		((ImageCell)cell.body).Image = (System.Drawing.Image)data;
            //	}
            //}

#if FORMULA
			if (formulaRanges.Count > 0)
			{
				// clear old references
				ClearCellReferenceList(cell);
			}

			// clear cell formula
			cell.InnerFormula = null;

			// clear formula status
			cell.formulaStatus = FormulaStatus.Normal;

#endif // FORMULA

            UpdateCellData(cell, data);
        }

        /// <summary>
        ///     Update data for cell without doing any format and formula evalution.
        /// </summary>
        /// <param name="cell">cell to be updated</param>
        /// <param name="data">data to be updated</param>
        /// <param name="dirtyCellStack">A stack to save cells that are marked as dirty cell, the dirty cell will be updated delay</param>
        internal void UpdateCellData(Cell cell, object data, Stack<List<Cell>> dirtyCellStack = null)
        {
            if (cell.body != null) data = cell.body.OnSetData(data);

            cell.InnerData = data;

            if (HasSettings(WorksheetSettings.Edit_AutoFormatCell))
                DataFormatterManager.Instance.FormatCell(cell);
            else
                cell.InnerDisplay = Convert.ToString(data);

#if WPF
            cell.formattedText = null;

            //if (cell.FormattedText == null || cell.FormattedText.Text != cell.InnerDisplay)
            //{
            //	float fontSize = cell.InnerStyle.FontSize * this.scaleFactor * (96f / 72f);

            //	cell.FormattedText = new System.Windows.Media.FormattedText(cell.InnerDisplay, 
            //		System.Globalization.CultureInfo.CurrentCulture,
            //		System.Windows.FlowDirection.LeftToRight,
            //		ResourcePoolManager.Instance.GetTypeface(cell.InnerStyle.FontName),
            //		fontSize,
            //		ResourcePoolManager.Instance.GetBrush(cell.InnerStyle.TextColor));
            //}
#endif

            AfterCellDataUpdate(cell, dirtyCellStack);
        }

        internal bool viewDirty;

        private void AfterCellDataUpdate(Cell cell, Stack<List<Cell>> dirtyCellStack = null)
        {
#if FORMULA
			if ((this.settings & WorksheetSettings.Formula_AutoUpdateReferenceCell)
				== WorksheetSettings.Formula_AutoUpdateReferenceCell)
			{
				UpdateReferencedFormulaCells(cell, dirtyCellStack);
			}
#endif // FORMULA

#if DRAWING
			if (cell.Data is Drawing.RichText)
			{
				var rt = (Drawing.RichText)cell.Data;

				rt.SuspendUpdateText();
				rt.Size = cell.Bounds.Size;
				rt.TextWrap = cell.InnerStyle.TextWrapMode;
				rt.ResumeUpdateText();
				rt.UpdateText();
			}
			else
#endif // DRAWING
            {
                cell.FontDirty = true;
            }

            if (controlAdapter != null
                && !viewDirty
                && !suspendDataChangedEvent)
                RequestInvalidate();

            if (!suspendDataChangedEvent)
            {
                var header = cols[cell.Column];

                if (header.Body != null) header.Body.OnDataChange(cell.Row, cell.Row);

                // raise text changed event
                RaiseCellDataChangedEvent(cell);
            }
        }

        #endregion // Set Data

        #region Set/Remove CellBody

        /// <summary>
        ///     Set body of cell at specified position of grid
        /// </summary>
        /// <param name="row">number of row</param>
        /// <param name="col">number of column</param>
        /// <param name="body">body to be set</param>
        public void SetCellBody(int row, int col, ICellBody body)
        {
            if (row < 0) throw new ArgumentOutOfRangeException("row");
            if (row >= cells.RowCapacity) throw new ArgumentOutOfRangeException("row");
            if (col < 0) throw new ArgumentOutOfRangeException("col");
            if (col >= cells.ColCapacity) throw new ArgumentOutOfRangeException("col");

            var cell = cells[row, col];

            if (cell == null)
            {
                if (body == null)
                    return;
                cell = CreateCell(row, col);
            }

            SetCellBody(cell, body);
        }

        public void SetCellBody(CellPosition pos, ICellBody body)
        {
            pos = FixPos(pos);
            SetCellBody(pos.Row, pos.Col, body);
        }

        public void SetCellBody(string address, ICellBody body)
        {
            if (!CellPosition.IsValidAddress(address))
                throw new InvalidAddressException(address);

            SetCellBody(new CellPosition(address), body);
        }

        /// <summary>
        ///     Set body of cell into specified row
        /// </summary>
        /// <param name="cell">cell to be set</param>
        /// <param name="body">body to be set</param>
        internal void SetCellBody(Cell cell, ICellBody body)
        {
            cell.Body = body;

            //if (body != null)
            //{
            //	body.OnSetup(cell);

            //	// why?
            //	UpdateCellFont(cell);
            //}

            //cell.UpdateContentBounds();

            RequestInvalidate();
        }

        /// <summary>
        ///     Remove cell body from specified cell
        /// </summary>
        /// <param name="pos">position of specified cell</param>
        public void RemoveCellBody(CellPosition pos)
        {
            RemoveCellBody(pos.Row, pos.Col);
        }

        /// <summary>
        ///     Remove cell body from specified cell
        /// </summary>
        /// <param name="row">number of row</param>
        /// <param name="col">number of column</param>
        public void RemoveCellBody(int row, int col)
        {
            var cell = cells[row, col];
            if (cell != null)
            {
                cell.body = null;
                RequestInvalidate();
            }
        }

        #endregion // Set/Remove CellBody

        #region Get Data/Text

        /// <summary>
        ///     Get cell data from specified address or name
        /// </summary>
        /// <param name="addressOrName">address or name to locate a cell</param>
        /// <returns>data from cell</returns>
        /// <exception cref="InvalidAddressException">throw if specified address or name is invalid</exception>
        public object GetCellData(string addressOrName)
        {
            if (CellPosition.IsValidAddress(addressOrName)) return GetCellData(new CellPosition(addressOrName));

            NamedRange range;
            if (registeredNamedRanges.TryGetValue(addressOrName, out range)) return GetCellData(range.StartPos);

            throw new InvalidAddressException(addressOrName);
        }

        /// <summary>
        ///     Get data from specified cell
        /// </summary>
        /// <param name="pos">Position of cell to get data</param>
        /// <returns>Data of cell</returns>
        public object GetCellData(CellPosition pos)
        {
            return GetCellData(pos.Row, pos.Col);
        }

        /// <summary>
        ///     Get data from specified cell
        /// </summary>
        /// <param name="row">zero-based number of row</param>
        /// <param name="col">zero-based number of column</param>
        /// <returns>data of cell</returns>
        public object GetCellData(int row, int col)
        {
            if (row < 0 || row >= rows.Count) return null;
            if (col < 0 || col >= cols.Count) return null;

            var cell = cells[row, col];
            return cell == null ? null : cell.InnerData;
        }

        /// <summary>
        ///     Get data from specified cell
        /// </summary>
        /// <param name="addressOrName">address or name to locate the cell</param>
        /// <returns>data of cell</returns>
        /// <exception cref="InvalidAddressException">throw if specified address or name is invalid</exception>
        public T GetCellData<T>(string addressOrName)
        {
            if (CellPosition.IsValidAddress(addressOrName)) return GetCellData<T>(new CellPosition(addressOrName));

            NamedRange range;
            if (registeredNamedRanges.TryGetValue(addressOrName, out range)) return GetCellData<T>(range.StartPos);

            throw new InvalidAddressException(addressOrName);
        }

        /// <summary>
        ///     Get data from specified cell
        /// </summary>
        /// <param name="pos">position to locate the cell</param>
        /// <returns>data of cell</returns>
        public T GetCellData<T>(CellPosition pos)
        {
            return CellUtility.ConvertData<T>(GetCellData(pos));
        }

        /// <summary>
        ///     Get data in specified type from a cell
        /// </summary>
        /// <typeparam name="T">type of data will be converted into</typeparam>
        /// <param name="row">number of row to locate a cell</param>
        /// <param name="col">number of column to locate a cell</param>
        /// <returns></returns>
        public T GetCellData<T>(int row, int col)
        {
            return CellUtility.ConvertData<T>(GetCellData(row, col));
        }

        /// <summary>
        ///     Try get number data from cell at specified position. If the data is string,
        ///     this method will try to convert the string into number value.
        /// </summary>
        /// <param name="row">Number of row of the cell to get data.</param>
        /// <param name="col">Number of column of the cell to get data.</param>
        /// <param name="val">Number data returned and converted from cell.</param>
        /// <returns>True if data can be get and converted; Otherwise return false.</returns>
        public bool TryGetNumberData(int row, int col, out double val)
        {
            var cell = cells[row, col];

            if (cell == null)
            {
                val = 0;
                return false;
            }

            return CellUtility.TryGetNumberData(cell.Data, out val);
        }

        /// <summary>
        ///     Get cell display text by specified address
        /// </summary>
        /// <param name="address">address to locate a cell</param>
        /// <returns>display text in string returned from specified cell</returns>
        public string GetCellText(string address)
        {
            if (CellPosition.IsValidAddress(address)) return GetCellText(new CellPosition(address));

            if (RangePosition.IsValidAddress(address)) return GetCellText(new RangePosition(address).StartPos);

            NamedRange range;
            if (NamedRange.IsValidName(address)
                && TryGetNamedRange(address, out range))
                return GetCellText(range.StartPos);
            throw new InvalidAddressException(address);
        }

        /// <summary>
        ///     Get formatted cell text from spcified position
        /// </summary>
        /// <param name="pos">position to be get</param>
        /// <returns>formatted cell's text</returns>
        public string GetCellText(CellPosition pos)
        {
            return GetCellText(pos.Row, pos.Col);
        }

        /// <summary>
        ///     Get formatted cell text from specified position
        /// </summary>
        /// <param name="row">Row of position</param>
        /// <param name="col">Col of position</param>
        /// <returns>Text of cell</returns>
        public string GetCellText(int row, int col)
        {
            var cell = cells[row, col];
            if (cell == null) return string.Empty;
            return string.IsNullOrEmpty(cell.DisplayText) ? string.Empty : cell.DisplayText;
        }

        #endregion // Get Data/Text
    }

    #region Cell

    /// <summary>
    ///     Represents the cell on worksheet. Cell instances are completely managed by ReoGrid.
    ///     To create custom cell, use <code>CellBody</code> class or </code>ICellBody</code> interface instead.
    /// </summary>
    /// <seealso cref="CellBody" />
    /// <seealso cref="ICellBody" />
    [Serializable]
    public partial class Cell //: IFormulaSource
    {
        #region Constructor

        /// <summary>
        ///     Instance of worksheet that the cell belongs to
        /// </summary>
        [NonSerialized] private Worksheet worksheet;

        /// <summary>
        ///     Get the worksheet that is the owner of this cell
        /// </summary>
        public Worksheet Worksheet
        {
            get { return worksheet; }
        }

        /// <summary>
        ///     Construct cell instance with specified owner grid instance.
        ///     Cell instance managed by ReoGrid core, it cannot be created by user-code.
        /// </summary>
        /// <param name="worksheet">Owner grid control for this cell</param>
        internal Cell(Worksheet worksheet)
        {
            this.worksheet = worksheet;
            FontDirty = true;
        }

        #endregion // Constructor

        #region Position

        internal CellPosition InternalPos;

        internal int InternalRow
        {
            get { return InternalPos.Row; }
            set { InternalPos.Row = value; }
        }

        internal int InternalCol
        {
            get { return InternalPos.Col; }
            set { InternalPos.Col = value; }
        }

        /// <summary>
        ///     Get zero-based row index of this cell
        /// </summary>
        public int Row
        {
            get { return InternalPos.Row; }
        }

        /// <summary>
        ///     Get zero-based column index of this cell
        /// </summary>
        public int Column
        {
            get { return InternalPos.Col; }
        }

        /// <summary>
        ///     Get position of this cell.
        /// </summary>
        public CellPosition Position
        {
            get { return InternalPos; }
        }

        /// <summary>
        ///     Get position as a range position which contains the cell rowspan and colspan.
        ///     This property is useful when the cell is a merged cell.
        /// </summary>
        public RangePosition PositionAsRange
        {
            get { return new RangePosition(InternalRow, InternalCol, rowspan, colspan); }
        }

        /// <summary>
        ///     Get address string of this cell
        /// </summary>
        public string Address
        {
            get { return RGUtility.ToAddress(InternalPos.Row, InternalPos.Col); }
        }
        //internal ReferenceRange ReferenceRange { get; set; }

        //public ReferenceRange GetContainedRange(){return this.ReferenceRange;}

        #endregion // Postion

        #region Rowspan & Colspan

        private short colspan;

        internal short Colspan
        {
            get { return colspan; }
            set { colspan = value; }
        }

        private short rowspan;

        internal short Rowspan
        {
            get { return rowspan; }
            set { rowspan = value; }
        }

        /// <summary>
        ///     Get number of colspan
        /// </summary>
        /// <returns></returns>
        public short GetColspan()
        {
            return colspan;
        }

        /// <summary>
        ///     Get number of rowspan
        /// </summary>
        /// <returns></returns>
        public short GetRowspan()
        {
            return rowspan;
        }

        #endregion // Rowspan & Colspan

        #region Location & Size

        [NonSerialized] private Rectangle bounds;

        internal Rectangle Bounds
        {
            get { return bounds; }
            set { bounds = value; }
        }

        internal double Width
        {
            get { return bounds.Width; }
            set { bounds.Width = value; }
        }

        internal double Height
        {
            get { return bounds.Height; }
            set { bounds.Height = value; }
        }

        internal double Top
        {
            get { return bounds.Y; }
            set { bounds.Y = value; }
        }

        internal double Left
        {
            get { return bounds.X; }
            set { bounds.X = value; }
        }

        internal double Right
        {
            get { return bounds.Right; }
            set { bounds.Width += bounds.Right - value; }
        }

        internal double Bottom
        {
            get { return bounds.Bottom; }
            set { bounds.Height += bounds.Bottom - value; }
        }

        #endregion // Location & Size

        #region Data Format

        private CellDataFormatFlag dataFormat;

        /// <summary>
        ///     Get or set the data format type
        /// </summary>
        public CellDataFormatFlag DataFormat
        {
            get { return dataFormat; }
            set { dataFormat = value; }
        }

        private object dataFormatArgs;

        /// <summary>
        ///     Get or set the argument of data format type
        /// </summary>
        public object DataFormatArgs
        {
            get { return dataFormatArgs; }
            set { dataFormatArgs = value; }
        }

        #endregion // Data Format

        #region Data, Display

        /// <summary>
        ///     cell data
        /// </summary>
        internal object InnerData { get; set; }

        /// <summary>
        ///     Get or set cell data
        /// </summary>
        public object Data
        {
            get { return InnerData; }
            set
            {
                if (worksheet != null)
                    // update cell data
                    worksheet.SetSingleCellData(this, value);
                else
                    InnerData = value;
            }
        }

        /// <summary>
        ///     Get and convert data into specified type
        /// </summary>
        /// <typeparam name="T">Type try to convert</typeparam>
        /// <returns>Converted data in specified type</returns>
        public T GetData<T>()
        {
            return CellUtility.ConvertData<T>(InnerData);
        }

        /// <summary>
        ///     cell formula
        /// </summary>
        internal string InnerFormula { get; set; }
#if FORMULA
		internal FormulaStatus formulaStatus;

		/// <summary>
		/// Determine the status of formula calculation
		/// </summary>
		public FormulaStatus FormulaStatus
		{
			get { return this.formulaStatus; }
		}
#endif // FORMULA

        /// <summary>
        ///     cell display text.
        /// </summary>
        internal string InnerDisplay { get; set; }

        /// <summary>
        ///     Get the text displayed in cell. Set <code>Data</code> property to change this text.
        /// </summary>
        public string DisplayText
        {
            get { return InnerDisplay; }
        }

        /// <summary>
        ///     Determine whether or not allow to change data of this cell.
        /// </summary>
        public bool IsReadOnly { get; set; }

        /// <summary>
        ///     Start edit this cell.
        /// </summary>
        public void StartEdit()
        {
            ValidateAssociation();

            worksheet.StartEdit(this);
        }

        /// <summary>
        ///     Finish edit this cell.
        /// </summary>
        public void EndEdit()
        {
            ValidateAssociation();

            if (worksheet.CurrentEditingCell == this) worksheet.EndEdit(this);
        }

        /// <summary>
        ///     Expand or shrink row height to fit the display text in this cell.
        /// </summary>
        public void ExpandRowHeight()
        {
            ValidateAssociation();
            worksheet.ExpandRowHeightToFitCell(this);
        }

        /// <summary>
        ///     Expand or shrink column width to fit the display text in this cell.
        /// </summary>
        public void ExpandColumnWidth()
        {
            ValidateAssociation();
            worksheet.ExpandColumnWidthFitToCell(this);
        }

        private void ValidateAssociation()
        {
            if (worksheet == null)
                throw new ReferenceObjectNotAssociatedException("Cell not associated to any worksheet.");
        }

        #endregion // Data and Display

        #region Formula

#if FORMULA
		[NonSerialized]
		internal STNode formulaTree;

		internal STNode FormulaTree { get { return this.formulaTree; } set { this.formulaTree
 = value; } }
#endif

        /// <summary>
        ///     Get or set the cell formula
        /// </summary>
        /// <returns>Formula of cell in string</returns>
        public string Formula
        {
            get { return InnerFormula; }
            set
            {
                if (InnerFormula != value)
                {
                    if (Worksheet == null)
                    {
                        InnerFormula = value;
                        InnerData = null;
                        InnerDisplay = null;
#if DEBUG
                        Logger.Log("cell", "Cell not attached to any worksheet: {0}", Position.ToAddress());
#endif // DEBUG
                    }
                    else
                    {
#if FORMULA
						if (string.IsNullOrEmpty(value))
						{
							this.worksheet.DeleteCellFormula(this);
						}
						else
						{
							this.worksheet.SetCellFormula(this, value);
							this.worksheet.RecalcCell(this);
						}
#else // FORMULA
                        worksheet.SetSingleCellData(this, value);
#endif // FORMULA
                    }
                }
            }
        }

        /// <summary>
        ///     Identify whether or not this cell contains the formula and the value has been evaluated
        /// </summary>
        public bool HasFormula
        {
            get { return !string.IsNullOrEmpty(InnerFormula); }
        }

        #endregion // Formula

        #region Style

        internal WorksheetRangeStyle InnerStyle { get; set; }

        internal StyleParentKind StyleParentKind { get; set; }

#if DEBUG
        private static int _count;
#endif // DEBUG

        internal void CreateOwnStyle()
        {
            InnerStyle = new WorksheetRangeStyle(InnerStyle);

            StyleParentKind = StyleParentKind.Own;

#if DEBUG
            if (++_count % 50 == 0) Logger.Log("style", "new style created, count: " + _count);
#endif // DEBUG
        }

        [NonSerialized] private ReferenceCellStyle referenceStyle;

        /// <summary>
        ///     Get or set style object of cell.
        /// </summary>
        public ReferenceCellStyle Style
        {
            get
            {
                if (referenceStyle == null) referenceStyle = new ReferenceCellStyle(this);

                return referenceStyle;
            }
            set
            {
                if (value == null)
                {
                    Worksheet.RemoveRangeStyles(new RangePosition(InternalPos, InternalPos), PlainStyleFlag.All);
                }
                else
                {
                    var anotherCell = value.Cell;
                    if (anotherCell == null) return;

                    var anotherStyle = anotherCell.InnerStyle;

                    switch (anotherCell.StyleParentKind)
                    {
                        case StyleParentKind.Col:
                        case StyleParentKind.Row:
                        case StyleParentKind.Root:
                        case StyleParentKind.Range:
                            InnerStyle = anotherCell.InnerStyle;
                            StyleParentKind = anotherCell.StyleParentKind;
                            break;

                        default:
                        case StyleParentKind.Own:
                            InnerStyle = new WorksheetRangeStyle(anotherCell.InnerStyle);
                            StyleParentKind = StyleParentKind.Own;
                            break;
                    }
                }
            }
        }

        /// <summary>
        ///     Checks whether or not this cell is visible. (Cells on hidden rows or columns will become invisibility)
        /// </summary>
        public bool IsVisible
        {
            get
            {
                if (Worksheet == null) return false;

                return Worksheet.IsCellVisible(this);
            }
        }

        // todo: support multi-lines
        private double distributedIndentSpacing;

        internal double DistributedIndentSpacing
        {
            get { return distributedIndentSpacing; }
            set { distributedIndentSpacing = value; }
        }

        private double distributedIndentSpacingPrint;

        internal double DistributedIndentSpacingPrint
        {
            get { return distributedIndentSpacingPrint; }
            set { distributedIndentSpacingPrint = value; }
        }

        #endregion // Style

        #region Border Wraps

        private CellBorderProperty borderProperty;

        /// <summary>
        ///     Get border property from cell.
        /// </summary>
        public CellBorderProperty Border
        {
            get
            {
                if (borderProperty == null) borderProperty = new CellBorderProperty(this);

                return borderProperty;
            }
        }

        #endregion // Border Wraps

        #region Cell Body

        internal void UpdateContentBounds()
        {
            if (body != null)
            {
                var cb = new Rectangle(InnerStyle.Padding.Left, InnerStyle.Padding.Top,
                    bounds.Width - 1 - InnerStyle.Padding.Left - InnerStyle.Padding.Right,
                    bounds.Height - 1 - InnerStyle.Padding.Top - InnerStyle.Padding.Bottom);

                if (body.Bounds != cb)
                {
                    body.Bounds = cb;
                    body.OnBoundsChanged();
                }
            }
        }

        /// <summary>
        ///     Get or set the user data attaching to this cell.
        /// </summary>
        public object Tag { get; set; }

        internal ICellBody body;

        /// <summary>
        ///     Get or set the cell body.
        /// </summary>
        public ICellBody Body
        {
            get { return body; }
            set
            {
                if (body != value)
                {
                    body = value;

                    if (body != null)
                    {
                        body.OnSetup(this);

                        if (body is CellBody) ((CellBody)body).InnerCell = this;
                    }

                    UpdateContentBounds();

                    if (Worksheet != null) Worksheet.RequestInvalidate();
                }
            }
        }

        #endregion

        #region Utility

        /// <summary>
        ///     Clone this cell instance by copying its all of properties
        /// </summary>
        /// <returns>new cell instance copied from this cell</returns>
        public Cell Clone()
        {
            var cell = new Cell(Worksheet);
            CellUtility.CopyCell(cell, this);
            return cell;
        }

        public override string ToString()
        {
            return "Cell[" + Address + "]";
        }

        #endregion // Utility
    }

    #endregion

    #region CellElementFlag

    /// <summary>
    ///     Cell element flags
    /// </summary>
    [Flags]
    public enum CellElementFlag
    {
        /// <summary>
        ///     All elements
        /// </summary>
        All = Data | Formula | DataFormat | Style | Border | Body,

        /// <summary>
        ///     Cell value
        /// </summary>
        Data = 0x1,

        /// <summary>
        ///     Cell formuals
        /// </summary>
        Formula = 0x2,

        /// <summary>
        ///     Cell body
        /// </summary>
        Body = 0x4,

        /// <summary>
        ///     Data format setting
        /// </summary>
        DataFormat = 0x8,

        /// <summary>
        ///     Styles
        /// </summary>
        Style = 0x10,

        /// <summary>
        ///     Border around the cell
        /// </summary>
        Border = 0x20
    }

    #endregion // CellElementFlag
}

namespace unvell.ReoGrid.Utility
{
    #region Cell Utility

    /// <summary>
    ///     Common utility functions for cell.
    /// </summary>
    public static class CellUtility
    {
        /// <summary>
        ///     Clone new cell instance from another cell.
        /// </summary>
        /// <param name="toCell">Target cell to be copied into.</param>
        /// <param name="fromCell">Original cell to be copied from.</param>
        /// <returns>Instance of cell cloned</returns>
        public static void CopyCell(Cell toCell, Cell fromCell)
        {
            // base
            toCell.InternalPos = fromCell.InternalPos;
            toCell.Rowspan = fromCell.Rowspan;
            toCell.Colspan = fromCell.Colspan;
            toCell.MergeStartPos = fromCell.MergeStartPos;
            toCell.MergeEndPos = fromCell.MergeEndPos;
            toCell.Bounds = fromCell.Bounds;

            // content
            CopyCellContent(toCell, fromCell);
        }

        /// <summary>
        ///     Copy all content from a cell to another cell.
        /// </summary>
        /// <param name="toCell">Target cell to be copied into.</param>
        /// <param name="fromCell">Source cell to be copied from.</param>
        public static void CopyCellContent(Cell toCell, Cell fromCell)
        {
            // style & render
            toCell.InnerStyle = new WorksheetRangeStyle(fromCell.InnerStyle);
            toCell.StyleParentKind = fromCell.StyleParentKind;
            toCell.TextBounds = fromCell.TextBounds;
            toCell.RenderHorAlign = fromCell.RenderHorAlign;
            toCell.RenderColor = fromCell.RenderColor;
            toCell.DistributedIndentSpacing = fromCell.DistributedIndentSpacing;

#if WINFORM
			toCell.RenderFont = fromCell.RenderFont;
#endif // WINFORM

            // data format
            toCell.DataFormat = fromCell.DataFormat;
            toCell.DataFormatArgs = fromCell.DataFormatArgs;
            toCell.Body = fromCell.body == null ? null : fromCell.body.Clone();

            // cell data
            toCell.InnerData = fromCell.InnerData;
            toCell.InnerDisplay = fromCell.DisplayText;

#if FORMULA
			toCell.Formula = fromCell.Formula;

			if (fromCell.formulaTree != null)
			{
				toCell.formulaTree = (STNode)fromCell.formulaTree.Clone();
			}

			// do not copy formula status since SetPartialGrid method will rebuilt formula status.
			//toCell.formulaStatus = fromCell.formulaStatus;
#endif // FORMULA

            // properties
            toCell.FontDirty = fromCell.FontDirty;
            toCell.IsReadOnly = fromCell.IsReadOnly;

            // custom content
            toCell.Tag = fromCell.Tag;
        }

        /// <summary>
        ///     Check specified data from a cell is whether or not a number.
        /// </summary>
        /// <param name="cell">Cell instance of data.</param>
        /// <returns>True if data from the cell is number.</returns>
        public static bool IsNumberData(Cell cell)
        {
            var data = cell.InnerData;
            return data is double || data is int || data is float || data is long || data is byte || data is decimal;
        }

        /// <summary>
        ///     Check specified data is whether or not a number.
        /// </summary>
        /// <param name="data">Data to be check.</param>
        /// <returns>True if data is number.</returns>
        public static bool IsNumberData(object data)
        {
            return data is double || data is int || data is float || data is long || data is byte || data is decimal;
        }

        /// <summary>
        ///     Try get double value from specified cell.
        /// </summary>
        /// <param name="cell">Instance of cell.</param>
        /// <param name="value">The output value converted from data.</param>
        /// <returns>True if convert is succesful.</returns>
        public static bool TryGetNumberData(Cell cell, out double value)
        {
            return TryGetNumberData(cell.InnerData, out value);
        }

        /// <summary>
        ///     Try get double value from specified object.
        /// </summary>
        /// <param name="data">Data to be check and converted.</param>
        /// <param name="value">The output value converted from data.</param>
        /// <returns>True if convert is succesful.</returns>
        public static bool TryGetNumberData(object data, out double value)
        {
            if (data is double)
            {
                value = (double)data;
                return true;
            }

            if (data is int || data is long || data is float || data is byte || data is decimal)
            {
                value = (double)Convert.ChangeType(data, typeof(double));
                return true;
            }

            if (data is string) return double.TryParse((string)data, out value);

            value = 0;
            return false;
        }

        #region ConvertCellData

        /// <summary>
        ///     Convert cell data into another type if possible.
        /// </summary>
        /// <param name="data">Cell data to be converted.</param>
        /// <returns>Output object after convert.</returns>
        public static T ConvertData<T>(object data)
        {
            T value;
            ConvertData(data, out value);
            return value;
        }

        /// <summary>
        ///     Convert cell data into another type if possible.
        /// </summary>
        /// <typeparam name="T">Type after convert.</typeparam>
        /// <param name="data">Cell data to be converted.</param>
        /// <param name="value">Output object after convert.</param>
        /// <returns>True if convert is successful.</returns>
        public static bool ConvertData<T>(object data, out T value)
        {
            if (data == null)
            {
                value = default(T);
                return false;
            }

            var targetType = typeof(T);

            if (data is string)
            {
                if (targetType == typeof(string))
                {
                    value = (T)data;
                    return true;
                }

                double tmpVal;
                if (!double.TryParse((string)data, out tmpVal))
                {
                    value = default(T);
                    return false;
                }

                data = tmpVal;
            }
            else if (data is StringBuilder)
            {
                if (targetType == typeof(StringBuilder))
                {
                    value = (T)data;
                    return true;
                }

                double tmpVal;
                if (!double.TryParse(((StringBuilder)data).ToString(), out tmpVal))
                {
                    value = default(T);
                    return false;
                }

                data = tmpVal;
            }

            // require string, but data is not in string, return false
            if (targetType != typeof(string) && (data is string || data is StringBuilder))
            {
                value = default(T);
                return false;
            }

            if (targetType == typeof(double))
                if (data is char)
                    data = (int)(char)data;

            value = (T)Convert.ChangeType(data, typeof(T));
            return true;
        }

        #endregion // ConvertCellData
    }

    #endregion // Cell Utility
}