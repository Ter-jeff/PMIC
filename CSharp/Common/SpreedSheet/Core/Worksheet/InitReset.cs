#define WPF

using SpreedSheet.Core;
using SpreedSheet.Interface;
using SpreedSheet.View.Controllers;
#if DEBUG
using System.Diagnostics;
#endif // DEBUG

#if WINFORM || WPF
//using CellArray = unvell.ReoGrid.Data.JaggedTreeArray<unvell.ReoGrid.ReoGridCell>;
//using HBorderArray = unvell.ReoGrid.Data.JaggedTreeArray<unvell.ReoGrid.Core.ReoGridHBorder>;
//using VBorderArray = unvell.ReoGrid.Data.JaggedTreeArray<unvell.ReoGrid.Core.ReoGridVBorder>;
using CellArray = unvell.ReoGrid.Data.Index4DArray<unvell.ReoGrid.Cell>;
using HBorderArray = unvell.ReoGrid.Data.Index4DArray<unvell.ReoGrid.Core.ReoGridHBorder>;
using VBorderArray = unvell.ReoGrid.Data.Index4DArray<unvell.ReoGrid.Core.ReoGridVBorder>;

#elif ANDROID || iOS
using CellArray = unvell.ReoGrid.Data.ReoGridCellArray;
using HBorderArray = unvell.ReoGrid.Data.ReoGridHBorderArray;
using VBorderArray = unvell.ReoGrid.Data.ReoGridVBorderArray;
#endif // ANDROID

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        private void InitGrid()
        {
            InitGrid(DefaultRows, DefaultCols);
        }

        private void InitGrid(int rows, int cols)
        {
#if DEBUG
            var sw = Stopwatch.StartNew();
            Debug.WriteLine("start creating worksheet...");
#endif // DEBUG

            SuspendUIUpdates();

            // resize spreadsheet to specified size
            Resize(rows, cols);

            if (controlAdapter != null)
            {
                renderScaleFactor = _scaleFactor + controlAdapter.BaseScale;

                var scv = ViewportController as IScalableViewportController;

                if (scv != null) scv.ScaleFactor = renderScaleFactor;
            }

            // restore root style
            RootStyle = new WorksheetRangeStyle(DefaultStyle);

            // initialize default settings
            settings = WorksheetSettings.Default;

            // reset selection 
            selectionRange = new RangePosition(0, 0, 1, 1);

#if PRINT
			// clear print settings
			if (this.printSettings != null) this.printSettings = null;
#endif // PRINT

#if DRAWING
			// drawing object
			this.drawingCanvas = new Drawing.WorksheetDrawingCanvas(this);
#endif // DRAWING

            ResumeUIUpdates();

            if (ViewportController != null)
                // reste viewport controller
                ViewportController.Reset();

#if EX_SCRIPT
			//settings |=
			//	// auto run script if loaded from file
			//		WorkbookSettings.Script_AutoRunOnload
			//	// confirm to user whether allow to run script after loaded from file
			//	| WorkbookSettings.Script_PromptBeforeAutoRun;

			//InitSRM();
			//this.worksheetObj = null;

			RaiseScriptEvent("onload");
#endif // EX_SCRIPT

#if DEBUG
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            if (ms > 10) Debug.WriteLine("creating worksheet done: " + ms + " ms.");
#endif // DEBUG
        }

        internal void Clear()
        {
            // hidden edit textbox
            EndEdit(EndEditReason.Cancel);

            // clear editing flag 
            endEditProcessing = false;

            // reset ActionManager
            if (controlAdapter != null)
            {
                var actionSupportedControl = controlAdapter.ControlInstance as IActionControl;

                if (actionSupportedControl != null) actionSupportedControl.ClearActionHistoryForWorksheet(this);
            }

#if OUTLINE
			// clear row outlines
			if (this.outlines != null)
			{
				ClearOutlines(RowOrColumn.Row | RowOrColumn.Column);
			}
#endif // OUTLINE

            // clear named ranges
            registeredNamedRanges.Clear();

            // clear highlight ranges
            if (highlightRanges != null) highlightRanges.Clear();

#if PRINT
			// clear page breaks
			if (this.pageBreakRows != null) this.pageBreakRows.Clear();
			if (this.pageBreakCols != null) this.pageBreakCols.Clear();

			if (this.userPageBreakRows != null) this.userPageBreakRows.Clear();
			if (this.userPageBreakCols != null) this.userPageBreakCols.Clear();

			this.printableRange = RangePosition.Empty;
			this.printSettings = null;
#endif // PRINT

#if DRAWING
			// drawing objects
			if (this.drawingCanvas != null)
			{
				this.drawingCanvas.Children.Clear();
			}
#endif // DRAWING

            // reset default width and height 
            DefaultColumnWidth = InitDefaultColumnWidth;
            defaultRowHeight = InitDefaultRowHeight;

            // clear root style
            RootStyle = new WorksheetRangeStyle(DefaultStyle);

            // clear focus highlight ranges
            FocusHighlightRange = null;

            // restore to default operation mode
            operationStatus = OperationStatus.Default;

            // restore settings
            settings = WorksheetSettings.Default;

            if (SettingsChanged != null) SettingsChanged(this, null);

#if FORMULA
			// clear formula referenced cells and ranges
			formulaRanges.Clear();

			// clear trace lines
			if (this.traceDependentArrows != null)
			{
				this.traceDependentArrows.Clear();
			}
#endif // FORMULA

#if EX_SCRIPT
			if (Srm != null)
			{
				RaiseScriptEvent("unload");
			}
#endif // EX_SCRIPT

            // unfreeze rows and columns
            var pos = FreezePos;
            if (pos.Row > 0 || pos.Col > 0) Unfreeze();

            if (ViewportController != null)
                // reset viewport controller
                ViewportController.Reset();

            // TODO: release objects inside cells and borders
            cells = new CellArray();
            hBorders = new HBorderArray();
            vBorders = new VBorderArray();

            // clear header & index
            rows.Clear();
            cols.Clear();

            // reset max row and column indexes
            maxRowHeader = -1;
            maxColumnHeader = -1;

            // reset highlight range color counter
            rangeHighlightColorCounter = 0;
        }

        /// <summary>
        ///     Reset control to default status.
        /// </summary>
        public void Reset()
        {
            Reset(DefaultRows, DefaultCols);
        }

        /// <summary>
        ///     Reset control and initialize to specified size
        /// </summary>
        /// <param name="rows">number of rows to be set after resting</param>
        /// <param name="cols">number of columns to be set after reseting</param>
        public void Reset(int rows, int cols)
        {
            // cancel editing mode
            EndEdit(EndEditReason.Cancel);

            // reset scale factor, need this?
            _scaleFactor = 1f;

            // clear grid
            Clear();

            // clear all actions belongs to this worksheet
            if (controlAdapter != null)
            {
                var actionSupportedControl = controlAdapter.ControlInstance as IActionControl;

                if (actionSupportedControl != null) actionSupportedControl.ClearActionHistoryForWorksheet(this);
            }

            // restore default cell size
            defaultRowHeight = InitDefaultRowHeight;
            DefaultColumnWidth = InitDefaultColumnWidth;

            // restore row header panel width
            _userRowHeaderWidth = false;

            // restore UI
            settings = WorksheetSettings.View_Default;

            // init grid
            InitGrid(rows, cols);

            // repaint
            RequestInvalidate();

            // raise reseting event
            if (Resetted != null) Resetted(this, null);
        }
    }
}