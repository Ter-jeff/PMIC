using SpreedSheet.View;
using unvell.ReoGrid;

namespace SpreedSheet.Rendering
{
    /// <summary>
    ///     Drawing context for rendering cells.
    /// </summary>
    public sealed class CellDrawingContext : DrawingContext
    {
        internal CellDrawingContext(Worksheet worksheet, DrawMode drawMode)
            : this(worksheet, drawMode, null)
        {
        }

        internal CellDrawingContext(Worksheet worksheet, DrawMode drawMode, IRenderer r)
            : base(worksheet, drawMode, r)
        {
            AllowCellClip = !worksheet.HasSettings(WorksheetSettings.View_AllowCellTextOverflow);
        }

        #region Cell Methods

        /// <summary>
        ///     Cell instance if enter a cell drawing event
        /// </summary>
        public Cell Cell { get; set; }

        internal bool AllowCellClip { get; set; }

        internal bool FullCellClip { get; set; }

        /// <summary>
        ///     Recall core renderer to draw cell text
        /// </summary>
        public void DrawCellText()
        {
            if (CurrentView is CellsViewport
                && Cell != null
                && !string.IsNullOrEmpty(Cell.DisplayText))
            {
                var view = (CellsViewport)CurrentView;

                var g = Graphics;

                var scaleFactor = Worksheet.renderScaleFactor;

                g.PopTransform();

                view.DrawCellText(this, Cell);

                g.PushTransform();
                if (scaleFactor != 1f) g.ScaleTransform(scaleFactor, scaleFactor);
                g.TranslateTransform(Cell.Left, Cell.Top);
            }
        }

        /// <summary>
        ///     Recall core renderer to draw cell background.
        /// </summary>
        public void DrawCellBackground()
        {
            if (CurrentView is CellsViewport
                && Cell != null)
            {
                var currentView = (CellsViewport)CurrentView;

                currentView.DrawCellBackground(this, Cell.InternalRow, Cell.InternalCol, Cell, true);
            }
        }

        #endregion // Cell Methods
    }
}