#define WPF

using System.Collections.Generic;
using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Reusable action group is one type of RGActionGroup to support repeat
    ///     operation to a specified range. It is good practice to make all reusable
    ///     action groups to inherit from this class.
    /// </summary>
    public class WorksheetReusableActionGroup : WorksheetReusableAction
    {
        private bool first = true;

        /// <summary>
        ///     Constructor of ReusableActionGroup
        /// </summary>
        /// <param name="range">Range to be appiled this action group</param>
        public WorksheetReusableActionGroup(RangePosition range)
            : base(range)
        {
            Actions = new List<WorksheetReusableAction>();
        }

        /// <summary>
        ///     Constructor of ReusableActionGroup
        /// </summary>
        /// <param name="range">Range to be appiled this action group</param>
        /// <param name="actions">Action list to be performed together</param>
        public WorksheetReusableActionGroup(RangePosition range, List<WorksheetReusableAction> actions)
            : base(range)
        {
            Actions = actions;
        }

        /// <summary>
        ///     All reusable actions stored in this list will be performed together.
        /// </summary>
        public List<WorksheetReusableAction> Actions { get; set; }

        /// <summary>
        ///     Do all actions stored in this action group
        /// </summary>
        public override void Do()
        {
            if (first)
            {
                for (var i = 0; i < Actions.Count; i++)
                    Actions[i].Worksheet = Worksheet;
                first = false;
            }

            for (var i = 0; i < Actions.Count; i++) Actions[i].Do();
        }

        /// <summary>
        ///     Undo all actions stored in this action group
        /// </summary>
        public override void Undo()
        {
            for (var i = Actions.Count - 1; i >= 0; i--) Actions[i].Undo();
        }

        /// <summary>
        ///     Get friendly name of this action group
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Multi-Aciton[" + Actions.Count + "]";
        }

        /// <summary>
        ///     Create cloned reusable action group from this action group
        /// </summary>
        /// <param name="range">Specified new range to apply this action group</param>
        /// <returns>New reusable action group cloned from this action group</returns>
        public override WorksheetReusableAction Clone(RangePosition range)
        {
            var clonedActions = new List<WorksheetReusableAction>();

            foreach (var action in Actions) clonedActions.Add(action.Clone(range));

            return new WorksheetReusableActionGroup(range, clonedActions);
        }
    }
}