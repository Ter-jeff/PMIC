#define WPF

using unvell.Common;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Base action for all actions that are used for worksheet operations.
    /// </summary>
    public abstract class BaseWorksheetAction : IUndoableAction
    {
        /// <summary>
        ///     Instance for the grid control will be setted before action performed.
        /// </summary>
        public Worksheet Worksheet { get; internal set; }

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
        ///     Get friendly name of this action.
        /// </summary>
        /// <returns>Get friendly name of this action.</returns>
        public abstract string GetName();
    }
}