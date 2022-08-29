#define WPF

using SpreedSheet.Core;

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Reusable action is one type of RGAction to support repeat operation
    ///     to a specified range. It is good practice to make all actions with
    ///     a range target to inherit from this class.
    /// </summary>
    public abstract class WorksheetReusableAction : BaseWorksheetAction
    {
        protected WorksheetReusableAction()
        {
        }

        /// <summary>
        ///     Constructor of RGReusableAction
        /// </summary>
        /// <param name="range">Range to be applied this action</param>
        public WorksheetReusableAction(RangePosition range)
        {
            Range = range;
        }

        /// <summary>
        ///     Range to be appiled this action
        /// </summary>
        public RangePosition Range { get; set; }

        /// <summary>
        ///     Create a copy from this action in order to apply the operation to another range.
        /// </summary>
        /// <param name="range">New range where this operation will be appiled to.</param>
        /// <returns>New action instance copied from this action.</returns>
        public abstract WorksheetReusableAction Clone(RangePosition range);
    }
}