#define WPF

using System.Collections.Generic;
using SpreedSheet.Core;
#if OUTLINE
using unvell.ReoGrid.Outline;
#endif // OUTLINE

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        /// <summary>
        ///     Clone this worksheet, create a new instance.
        /// </summary>
        /// <returns>New instance cloned from current worksheet.</returns>
        public Worksheet Clone(string newName = null)
        {
            if (workbook == null)
                throw new ReferenceObjectNotAssociatedException("worksheet must be added into workbook to do this");

            if (string.IsNullOrEmpty(newName)) newName = workbook.GetAvailableWorksheetName();

            var newSheet = new Worksheet(workbook, null, 0, 0)
            {
                name = newName,
                RootStyle = new WorksheetRangeStyle(RootStyle),
                //controlStyle = this.controlStyle,

                DefaultColumnWidth = DefaultColumnWidth,
                defaultRowHeight = defaultRowHeight,
                registeredNamedRanges = new Dictionary<string, NamedRange>(registeredNamedRanges),
#if OUTLINE
				outlines =
 this.outlines == null ? null : new Dictionary<RowOrColumn, OutlineCollection<ReoGridOutline>>(this.outlines),
#endif // OUTLINE
                highlightRanges = new List<HighlightRange>(highlightRanges),

#if FREEZE
#endif // FREEZE

#if PRINT
				pageBreakRows = this.pageBreakRows == null ? null : new List<int>(this.pageBreakRows),
				pageBreakCols = this.pageBreakCols == null ? null : new List<int>(this.pageBreakCols),
				userPageBreakCols = this.userPageBreakCols == null ? null : new List<int>(this.userPageBreakCols),
				userPageBreakRows = this.userPageBreakRows == null ? null : new List<int>(this.userPageBreakRows),
#endif // PRINT

                settings = settings
            };

            newSheet.rows.Capacity = rows.Count;
            newSheet.cols.Capacity = cols.Count;

            // copy headers

            foreach (var rheader in rows) newSheet.rows.Add(rheader.Clone(newSheet));

            foreach (var cheader in cols) newSheet.cols.Add(cheader.Clone(newSheet));

            // copy cells
            var partialGrid = GetPartialGrid(RangePosition.EntireRange);
            newSheet.SetPartialGrid(RangePosition.EntireRange, partialGrid);

            //this.IterateCells(ReoGridRange.EntireRange, (row, col, cell) =>
            //{
            //	var toCell = newSheet.CreateAndGetCell(row, col);
            //	ReoGridCellUtility.CopyCell(toCell, cell);

            //	return true;
            //});

            // copy drawing objects  (TODO: cloen all objects)
            //newSheet.drawingCanvas.Children.AddRange(this.drawingCanvas.Children);

            //var nvc = newSheet.viewportController as Views.NormalViewportController;
            //if (nvc != null)
            //{
            //	nvc.Bounds = this.viewportController.Bounds;
            //	newSheet.UpdateViewportControllBounds();
            //}

            // copy freeze info
            var frozenPos = FreezePos;

            if (frozenPos.Row > 0 || frozenPos.Col > 0) newSheet.FreezeToCell(frozenPos, FreezeArea);

            newSheet.ScaleFactor = ScaleFactor;

            newSheet.UpdateViewportController();

            return newSheet;
        }
    }
}