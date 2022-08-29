#define WPF

using System.Collections.Generic;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     The action group is one type of RGAction to support Do/Undo/Redo a series of actions.
    /// </summary>
    public class WorksheetActionGroup : BaseWorksheetAction
    {
        /// <summary>
        ///     Create instance for RGActionGroup
        /// </summary>
        public WorksheetActionGroup()
        {
            Actions = new List<BaseWorksheetAction>();
        }

        /// <summary>
        ///     Actions stored in this list will be Do/Undo/Redo together
        /// </summary>
        public List<BaseWorksheetAction> Actions { get; set; }

        /// <summary>
        ///     Do all actions stored in this action group
        /// </summary>
        public override void Do()
        {
            foreach (var action in Actions)
            {
                action.Worksheet = Worksheet;
                action.Do();
            }
        }

        /// <summary>
        ///     Undo all actions stored in this action group
        /// </summary>
        public override void Undo()
        {
            for (var i = Actions.Count - 1; i >= 0; i--)
            {
                var action = Actions[i];

                action.Worksheet = Worksheet;
                action.Undo();
            }
        }

        /// <summary>
        ///     Get friendly name of this action group
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "ReoGrid Action Group";
        }
    }
}