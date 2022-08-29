#define WPF

using SpreedSheet.Core;
using unvell.ReoGrid.Data;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action to create column filter
    /// </summary>
    public class CreateAutoFilterAction : BaseWorksheetAction
    {
        /// <summary>
        ///     Create action to create column filter
        /// </summary>
        /// <param name="range">filter range</param>
        public CreateAutoFilterAction(RangePosition range)
        {
            Range = range;
        }

        /// <summary>
        ///     Get filter apply range.
        /// </summary>
        public RangePosition Range { get; }

        /// <summary>
        ///     Get auto column filter instance created by this action. (Will be null before doing action)
        /// </summary>
        public AutoColumnFilter AutoColumnFilter { get; private set; }

        ///// <summary>
        ///// Create action to create column filter
        ///// </summary>
        ///// <param name="startColumn">zero-based number of column begin to create filter</param>
        ///// <param name="endColumn">zero-based number of column end to create filter</param>
        ///// <param name="titleRows">number of rows as title rows will not be included in filter and sort range</param>
        //public CreateAutoFilterAction(int startColumn, int endColumn, int titleRows = 1)
        //{
        //	this.StartColumn = startColumn;
        //	this.EndColumn = endColumn;
        //	this.TitleRows = titleRows;
        //}

        /// <summary>
        ///     Undo action to remove column filter that is created by this action
        /// </summary>
        public override void Undo()
        {
            if (AutoColumnFilter != null) AutoColumnFilter.Detach();
        }

        /// <summary>
        ///     Do action to create column filter
        /// </summary>
        public override void Do()
        {
            if (AutoColumnFilter == null)
                AutoColumnFilter = Worksheet.CreateColumnFilter(Range);
            else
                AutoColumnFilter.Attach(Worksheet);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns>friendly name of this action</returns>
        public override string GetName()
        {
            return "Create Column Filter";
        }
    }
}