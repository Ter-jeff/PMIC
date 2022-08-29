#define WPF

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action for removing worksheet
    /// </summary>
    public class RemoveWorksheetAction : WorkbookAction
    {
        /// <summary>
        ///     Create this action to insert worksheet
        /// </summary>
        /// <param name="index">Number of worksheet</param>
        /// <param name="worksheet">Worksheet instance</param>
        public RemoveWorksheetAction(int index, Worksheet worksheet)
        {
            Index = index;
            Worksheet = worksheet;
        }

        /// <summary>
        ///     Number of worksheet
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Worksheet instance
        /// </summary>
        public Worksheet Worksheet { get; }

        /// <summary>
        ///     Do this action to remove worksheet
        /// </summary>
        public override void Do()
        {
            Workbook.RemoveWorksheet(Index);
        }

        /// <summary>
        ///     Undo this action to restore the removed worksheet
        /// </summary>
        public override void Undo()
        {
            Workbook.InsertWorksheet(Index, Worksheet);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Remove Worksheet: " + Worksheet.Name;
        }
    }
}