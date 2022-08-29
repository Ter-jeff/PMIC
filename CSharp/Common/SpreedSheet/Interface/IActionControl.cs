using System;
using SpreedSheet.Core;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Events;

namespace SpreedSheet.Interface
{
    internal interface IActionControl
    {
        void DoAction(Worksheet sheet, BaseWorksheetAction action);
        void Undo();
        void Redo();
        void RepeatLastAction(RangePosition range);
        event EventHandler<WorkbookActionEventArgs> ActionPerformed;
        event EventHandler<WorkbookActionEventArgs> Undid;
        event EventHandler<WorkbookActionEventArgs> Redid;
        void ClearActionHistory();
        void ClearActionHistoryForWorksheet(Worksheet sheet);
    }
}