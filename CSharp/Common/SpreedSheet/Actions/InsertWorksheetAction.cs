#define WPF

namespace unvell.ReoGrid.Actions
{
    /// <summary>
    ///     Action for inserting worksheet
    /// </summary>
    public class InsertWorksheetAction : WorkbookAction
    {
        /// <summary>
        ///     Create this action to insert worksheet
        /// </summary>
        /// <param name="index">Number of worksheet</param>
        /// <param name="worksheet">Worksheet instance</param>
        public InsertWorksheetAction(int index, Worksheet worksheet)
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
        ///     Do this action to insert worksheet
        /// </summary>
        public override void Do()
        {
            Workbook.InsertWorksheet(Index, Worksheet);
        }

        /// <summary>
        ///     Undo this action to remove the inserted worksheet
        /// </summary>
        public override void Undo()
        {
            Workbook.RemoveWorksheet(Index);
        }

        /// <summary>
        ///     Get friendly name of this action
        /// </summary>
        /// <returns></returns>
        public override string GetName()
        {
            return "Insert Worksheet: " + Worksheet.Name;
        }
    }
}