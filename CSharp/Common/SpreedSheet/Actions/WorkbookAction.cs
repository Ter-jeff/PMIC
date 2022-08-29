#define WPF

using SpreedSheet.Core.Workbook;
using unvell.Common;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Represents an action of workbook.
    /// </summary>
    public abstract class WorkbookAction : IUndoableAction
    {
        /// <summary>
        ///     Create workbook action with specified workbook instance.
        /// </summary>
        /// <param name="workbook"></param>
        public WorkbookAction(IWorkbook workbook = null)
        {
            Workbook = workbook;
        }

        /// <summary>
        ///     Get the workbook instance.
        /// </summary>
        public IWorkbook Workbook { get; internal set; }

        /// <summary>
        ///     Do this action.
        /// </summary>
        public abstract void Do();

        /// <summary>
        ///     Undo this action.
        /// </summary>
        public abstract void Undo();

        /// <summary>
        ///     Redo this action.
        /// </summary>
        public virtual void Redo()
        {
            Do();
        }

        /// <summary>
        ///     Get the friendly name of this action.
        /// </summary>
        /// <returns></returns>
        public abstract string GetName();
    }
}