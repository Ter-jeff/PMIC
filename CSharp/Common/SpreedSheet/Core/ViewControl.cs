using System.Diagnostics;
using SpreedSheet.View.Controllers;

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        // reserved
        //private bool isLeadHeadHover = false;
        internal bool isLeadHeadSelected = false;

        /// <summary>
        ///     Check whether UI updates is suspending.
        /// </summary>
        public bool IsUIUpdatesSuspending { get; private set; }

        /// <summary>
        ///     Get or set viewport controller for worksheet.
        /// </summary>
        internal IViewportController ViewportController { get; set; }

        /// <summary>
        ///     Get or set view mode of current worksheet (Reserved)
        /// </summary>
        public ReoGridViewMode ViewMode { get; set; }

        /// <summary>
        ///     Suspend worksheet UI updates.
        /// </summary>
        public void SuspendUIUpdates()
        {
            IsUIUpdatesSuspending = true;
        }

        /// <summary>
        ///     Resume worksheet UI updates.
        /// </summary>
        public void ResumeUIUpdates()
        {
            if (IsUIUpdatesSuspending)
            {
                IsUIUpdatesSuspending = false;
                RequestInvalidate();
            }
        }

        internal void UpdateViewportController()
        {
            if (IsUIUpdatesSuspending) return;

#if DEBUG
            var sw = Stopwatch.StartNew();
#endif // DEBUG

            AutoAdjustRowHeaderPanelWidth();

            if (ViewportController != null)
            {
                ViewportController.UpdateController();

                if (IsFrozen) FreezeToCell(FreezePos, FreezeArea);
            }

#if PRINT // TODO: why do this here
			// update page breaks 
			if (this.HasSettings(WorksheetSettings.View_ShowPageBreaks)
				&& this.rows.Count > 0 && this.cols.Count > 0)
			{
				AutoSplitPage();
			}
#endif // PRINT

#if DEBUG
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            if (ms > 15) Debug.WriteLine("updating viewport controller takes " + sw.ElapsedMilliseconds + " ms.");
#endif // DEBUG
        }

        internal void UpdateViewportControlBounds()
        {
            // update boundary of viewportController 
            if (ViewportController != null && controlAdapter != null)
            {
                // don't compare Bounds before set
                ViewportController.Bounds = controlAdapter.GetContainerBounds();

                // need to update controller anytime when this method is called
                ViewportController.UpdateController();
            }
        }

        #region Invalidations

        /// <summary>
        ///     Request to repaint entire worksheet.
        /// </summary>
        public void RequestInvalidate()
        {
            if (!viewDirty && !IsUIUpdatesSuspending)
            {
                viewDirty = true;

                if (controlAdapter != null) controlAdapter.Invalidate();
            }
        }

        #endregion // Invalidations
    }
}