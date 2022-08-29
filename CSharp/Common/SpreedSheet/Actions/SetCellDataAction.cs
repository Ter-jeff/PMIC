#define WPF

using System;
using SpreedSheet.Core;
using unvell.ReoGrid.DataFormat;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Set data of cell action.
    /// </summary>
    public class SetCellDataAction : BaseWorksheetAction
    {
        private object backupData;

        //private string backupFormula;
        //private string displayBackup;
        private CellDataFormatFlag backupDataFormat;

        private object backupDataFormatArgs;

        //private Core.ReoGridRenderHorAlign backupRenderAlign;
        //private bool autoUpdateReferenceCells = false;
        private ushort? backupRowHeight = 0;

        //private bool isCellNull;

        /// <summary>
        ///     Create SetCellValueAction with specified index of row and column.
        /// </summary>
        /// <param name="row">index of row to set data.</param>
        /// <param name="col">index of column to set data.</param>
        /// <param name="data">data to be set.</param>
        public SetCellDataAction(int row, int col, object data)
        {
            Row = row;
            Col = col;
            Data = data;
        }

        /// <summary>
        ///     Create SetCellValueAction with specified index of row and column.
        /// </summary>
        /// <param name="pos">position to locate the cell to be set.</param>
        /// <param name="data">data to be set.</param>
        public SetCellDataAction(CellPosition pos, object data)
            : this(pos.Row, pos.Col, data)
        {
        }

        /// <summary>
        ///     Create action to set cell data.
        /// </summary>
        /// <param name="address">address to locate specified cell.</param>
        /// <param name="data">data to be set.</param>
        public SetCellDataAction(string address, object data)
        {
            var pos = new CellPosition(address);
            Row = pos.Row;
            Col = pos.Col;
            Data = data;
        }

        /// <summary>
        ///     Index of row to set data.
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        ///     Index of column to set data.
        /// </summary>
        public int Col { get; set; }

        /// <summary>
        ///     Data of cell.
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        ///     Do this operation.
        /// </summary>
        public override void Do()
        {
            var cell = Worksheet.CreateAndGetCell(Row, Col);

            backupData = cell.HasFormula ? "=" + cell.InnerFormula : cell.InnerData;

            backupDataFormat = cell.DataFormat;
            backupDataFormatArgs = cell.DataFormatArgs;

            try
            {
                Worksheet.SetSingleCellData(cell, Data);

                var rowHeightSettings = WorksheetSettings.Edit_AutoExpandRowHeight
                                        | WorksheetSettings.Edit_AllowAdjustRowHeight;

                var rowHeader = Worksheet.GetRowHeader(cell.InternalRow);

                backupRowHeight = rowHeader.InnerHeight;

                if ((Worksheet.settings & rowHeightSettings) == rowHeightSettings
                    && rowHeader.IsAutoHeight)
                    cell.ExpandRowHeight();
            }
            catch (Exception ex)
            {
                Worksheet.NotifyExceptionHappen(ex);
            }
        }

        public override void Redo()
        {
            Do();

            var cell = Worksheet.GetCell(Row, Col);

            if (cell != null)
                Worksheet.SelectRange(new RangePosition(cell.InternalRow, cell.InternalCol, cell.Rowspan,
                    cell.Colspan));
        }

        /// <summary>
        ///     Undo this operation.
        /// </summary>
        public override void Undo()
        {
            if (backupRowHeight != null)
            {
                var rowHeader = Worksheet.GetRowHeader(Row);

                if (rowHeader.InnerHeight != backupRowHeight) Worksheet.SetRowsHeight(Row, 1, (ushort)backupRowHeight);
            }

            var cell = Worksheet.GetCell(Row, Col);
            if (cell != null)
            {
                cell.DataFormat = backupDataFormat;
                cell.DataFormatArgs = backupDataFormatArgs;

                Worksheet.SetSingleCellData(cell, backupData);

                Worksheet.SelectRange(new RangePosition(cell.InternalRow, cell.InternalCol, cell.Rowspan,
                    cell.Colspan));
            }
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            var str = Data == null ? "null" : Data.ToString();
            return "Set Cell Value: " + (str.Length > 10 ? str.Substring(0, 7) + "..." : str);
        }
    }
}