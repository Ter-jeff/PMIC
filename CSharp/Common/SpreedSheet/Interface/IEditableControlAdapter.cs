using SpreedSheet.Core.Enum;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Interface
{
    internal interface IEditableControlAdapter
    {
        void ShowEditControl(Rectangle bounds, Cell cell);
        void HideEditControl();
        void SetEditControlText(string text);
        string GetEditControlText();
        void EditControlSelectAll();
        void SetEditControlCaretPos(int pos);
        int GetEditControlCaretPos();
        int GetEditControlCaretLine();
        void SetEditControlAlignment(GridHorAlign align);
        void EditControlApplySystemMouseDown();
        void EditControlCopy();
        void EditControlPaste();
        void EditControlCut();
        void EditControlUndo();
    }
}