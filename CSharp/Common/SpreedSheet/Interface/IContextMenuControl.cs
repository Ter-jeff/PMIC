using System.Windows.Controls;

namespace SpreedSheet.Interface
{
    internal interface IContextMenuControl
    {
        ContextMenu CellsContextMenu { get; }
        ContextMenu RowHeaderContextMenu { get; }
        ContextMenu ColumnHeaderContextMenu { get; }
        ContextMenu LeadHeaderContextMenu { get; }
    }
}