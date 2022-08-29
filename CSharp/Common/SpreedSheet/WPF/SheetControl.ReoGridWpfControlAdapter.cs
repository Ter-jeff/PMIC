using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SpreedSheet.Control;
using SpreedSheet.Core.Enum;
using SpreedSheet.Core.Workbook.Appearance;
using SpreedSheet.Interaction;
using SpreedSheet.Interface;
using SpreedSheet.Rendering;
using SpreedSheet.View;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;
using Point = unvell.ReoGrid.Graphics.Point;

namespace SpreedSheet.WPF
{
    public class SheetControlAdapter : IControlAdapter
    {
        #region Constructor

        public SheetControl _canvas;
        public InputTextBox EditTextBox { get; set; }

        internal SheetControlAdapter(SheetControl canvas)
        {
            _canvas = canvas;
            AddContextMenu();
        }

        #endregion // Constructor

        #region IControlAdapter Members

        public IVisualWorkbook ControlInstance
        {
            get { return _canvas; }
        }

        public ControlAppearanceStyle ControlStyle
        {
            get { return _canvas.controlStyle; }
        }

        public IRenderer Renderer
        {
            get { return _canvas.Renderer; }
        }

        public void ShowContextMenuStrip(ViewTypes viewType, Point containerLocation)
        {
            switch (viewType)
            {
                default:
                case ViewTypes.Cells:
                    _canvas.BaseContextMenu = _canvas.CellsContextMenu;
                    break;

                case ViewTypes.ColumnHeader:
                    _canvas.BaseContextMenu = _canvas.ColumnHeaderContextMenu;
                    break;

                case ViewTypes.RowHeader:
                    _canvas.BaseContextMenu = _canvas.RowHeaderContextMenu;
                    break;

                case ViewTypes.LeadHeader:
                    _canvas.BaseContextMenu = _canvas.LeadHeaderContextMenu;
                    break;
            }
        }

        public void AddContextMenu()
        {
            if (_canvas.CellsContextMenu == null)
            {
                var contextMenu = new ContextMenu();

                var cut = new MenuItem();
                cut.Header = "Cut";
                cut.Click += (sender, e) => _canvas.ActiveWorksheet.Cut();
                contextMenu.Items.Add(cut);

                var copy = new MenuItem();
                copy.Header = "Copy";
                copy.Click += (sender, e) => _canvas.ActiveWorksheet.Copy();
                contextMenu.Items.Add(copy);

                var paste = new MenuItem();
                paste.Header = "Paste";
                paste.Click += (sender, e) => _canvas.ActiveWorksheet.Paste();
                contextMenu.Items.Add(paste);

                contextMenu.Items.Add(new Separator());

                var selectAll = new MenuItem();
                selectAll.Header = "Select All";
                selectAll.Click += (sender, e) => _canvas.ActiveWorksheet.SelectAll();
                contextMenu.Items.Add(selectAll);

                contextMenu.Opened += (sender, e) =>
                {
                    cut.Visibility = true ? Visibility.Visible : Visibility.Collapsed;
                    copy.Visibility = true ? Visibility.Visible : Visibility.Collapsed;
                    paste.Visibility = Clipboard.ContainsText() ? Visibility.Visible : Visibility.Collapsed;
                    selectAll.Visibility = true ? Visibility.Visible : Visibility.Collapsed;
                };

                _canvas.CellsContextMenu = contextMenu;
            }
        }

        private Cursor _oldCursor;

        public void ChangeCursor(CursorStyle cursor)
        {
            _oldCursor = _canvas.Cursor;

            switch (cursor)
            {
                default:
                case CursorStyle.PlatformDefault:
                    _canvas.Cursor = Cursors.Arrow;
                    break;
                case CursorStyle.Selection:
                    _canvas.Cursor = _canvas.internalCurrentCursor;
                    break;
                case CursorStyle.Busy:
                    _canvas.Cursor = Cursors.AppStarting;
                    break;
                case CursorStyle.Hand:
                    _canvas.Cursor = Cursors.Hand;
                    break;
                case CursorStyle.FullColumnSelect:
                    _canvas.Cursor = _canvas.builtInFullColSelectCursor;
                    break;
                case CursorStyle.FullRowSelect:
                    _canvas.Cursor = _canvas.builtInFullRowSelectCursor;
                    break;
                case CursorStyle.ChangeRowHeight:
                    _canvas.Cursor = Cursors.SizeNS;
                    break;
                case CursorStyle.ChangeColumnWidth:
                    _canvas.Cursor = Cursors.SizeWE;
                    break;
                case CursorStyle.ResizeHorizontal:
                    _canvas.Cursor = Cursors.SizeWE;
                    break;
                case CursorStyle.ResizeVertical:
                    _canvas.Cursor = Cursors.SizeNS;
                    break;
                case CursorStyle.Move:
                    _canvas.Cursor = Cursors.SizeAll;
                    break;
                case CursorStyle.Cross:
                    _canvas.Cursor = _canvas.builtInCrossCursor;
                    break;
            }
        }

        public void RestoreCursor()
        {
            _canvas.Cursor = _oldCursor;
        }

        public void ChangeSelectionCursor(CursorStyle cursor)
        {
            switch (cursor)
            {
                default:
                case CursorStyle.PlatformDefault:
                    _canvas.internalCurrentCursor = Cursors.Arrow;
                    break;

                case CursorStyle.Hand:
                    _canvas.internalCurrentCursor = Cursors.Hand;
                    break;
            }
        }

        public Rectangle GetContainerBounds()
        {
            var w = _canvas.ActualWidth;
            var h = _canvas.ActualHeight + 1;

            if (_canvas._verScrollbar.Visibility == Visibility.Visible) w -= SheetControl.ScrollBarSize;

            if (_canvas._sheetTab.Visibility == Visibility.Visible
                || _canvas._horScrollbar.Visibility == Visibility.Visible)
                h -= SheetControl.ScrollBarSize;

            if (w < 0) w = 0;
            if (h < 0) h = 0;

            return new Rectangle(0, 0, w, h);
        }

        public void Focus()
        {
            _canvas.Focus();
        }

        public void Invalidate()
        {
            _canvas.InvalidateVisual();
        }

        public void ChangeBackColor(Color color)
        {
            _canvas.Background = new SolidColorBrush(color);
        }

        public bool IsVisible
        {
            get { return _canvas.Visibility == Visibility.Visible; }
            set { _canvas.Visibility = value ? Visibility.Visible : Visibility.Hidden; }
        }

        public Point PointToScreen(Point p)
        {
            return _canvas.PointToScreen(p);
        }

        public IGraphics PlatformGraphics
        {
            get { return null; }
        }

        public void ChangeBackgroundColor(SolidColor color)
        {
            _canvas.Background = new SolidColorBrush(color);
        }

        public void ShowTooltip(Point point, string content)
        {
            // not implemented
        }

        public ISheetTabControl SheetTabControl
        {
            get { return _canvas._sheetTab; }
        }

        public double BaseScale
        {
            get { return 0f; }
        }

        public double MinScale
        {
            get { return 0.1f; }
        }

        public double MaxScale
        {
            get { return 4f; }
        }

        #endregion // IControlAdapter Members

        #region IEditableControlInterface Members

        public void ShowEditControl(Rectangle bounds, Cell cell)
        {
            var sheet = _canvas.ActiveWorksheet;

            Color textColor;

            if (!cell.RenderColor.IsTransparent)
                textColor = cell.RenderColor;
            else if (cell.InnerStyle.HasStyle(PlainStyleFlag.TextColor))
                // cell text color, specified by SetRangeStyle
                textColor = cell.InnerStyle.TextColor;
            else
                // default cell text color
                textColor = _canvas.controlStyle[ControlAppearanceColors.GridText];

            Canvas.SetLeft(EditTextBox, bounds.X - 1);
            Canvas.SetTop(EditTextBox, bounds.Y);

            EditTextBox.Width = bounds.Width;
            EditTextBox.Height = bounds.Height;
            EditTextBox.RenderSize = bounds.Size;

            EditTextBox.CellSize = cell.Bounds.Size;
            EditTextBox.VAlign = cell.InnerStyle.VAlign;
            EditTextBox.FontFamily = new FontFamily(cell.InnerStyle.FontName);
            EditTextBox.FontSize = cell.InnerStyle.FontSize * sheet.ScaleFactor * 96f / 72f;
            EditTextBox.FontStyle = PlatformUtility.ToWPFFontStyle(cell.InnerStyle.fontStyles);
            EditTextBox.Foreground = Renderer.GetBrush(textColor);
            EditTextBox.Background = Renderer.GetBrush(cell.InnerStyle.HasStyle(PlainStyleFlag.BackColor)
                ? cell.InnerStyle.BackColor
                : _canvas.controlStyle[ControlAppearanceColors.GridBackground]);
            //EditTextBox.Background = Brushes.Transparent;
            EditTextBox.SelectionStart = EditTextBox.Text.Length;
            EditTextBox.TextWrap = cell.InnerStyle.TextWrapMode != TextWrapMode.NoWrap;
            EditTextBox.TextWrapping = cell.InnerStyle.TextWrapMode == TextWrapMode.NoWrap
                ? TextWrapping.NoWrap
                : TextWrapping.Wrap;

            EditTextBox.Visibility = Visibility.Visible;
            EditTextBox.Focus();
        }

        public void HideEditControl()
        {
            EditTextBox.Visibility = Visibility.Hidden;
        }

        public void SetEditControlText(string text)
        {
            EditTextBox.Text = text;
        }

        public string GetEditControlText()
        {
            return EditTextBox.Text;
        }

        public void EditControlSelectAll()
        {
            EditTextBox.SelectAll();
        }

        public void SetEditControlCaretPos(int pos)
        {
            EditTextBox.SelectionStart = pos;
        }

        public int GetEditControlCaretPos()
        {
            return EditTextBox.SelectionStart;
        }

        public int GetEditControlCaretLine()
        {
            return EditTextBox.GetLineIndexFromCharacterIndex(EditTextBox.SelectionStart);
        }

        public void SetEditControlAlignment(GridHorAlign align)
        {
            switch (align)
            {
                default:
                case GridHorAlign.Left:
                    EditTextBox.HorizontalAlignment = HorizontalAlignment.Left;
                    break;

                case GridHorAlign.Center:
                case GridHorAlign.DistributedIndent:
                    EditTextBox.HorizontalAlignment = HorizontalAlignment.Center;
                    break;

                case GridHorAlign.Right:
                    EditTextBox.HorizontalAlignment = HorizontalAlignment.Right;
                    break;
            }
        }

        public void EditControlApplySystemMouseDown()
        {
            Point p = Mouse.GetPosition(EditTextBox);

            p.X += 2; // fix 2 pixels (borders of left and right)
            p.Y -= 1; // fix 1 pixels (top)

            var caret = EditTextBox.GetCharacterIndexFromPoint(p, true);

            if (caret >= 0 && caret <= EditTextBox.Text.Length) EditTextBox.SelectionStart = caret;

            EditTextBox.Focus();
        }

        public void EditControlCopy()
        {
            EditTextBox.Copy();
        }

        public void EditControlPaste()
        {
            EditTextBox.Paste();
        }

        public void EditControlCut()
        {
            EditTextBox.Cut();
        }

        public void EditControlUndo()
        {
            EditTextBox.Undo();
        }

        #endregion

        #region IScrollableControlInterface Members

        public bool ScrollBarHorizontalVisible
        {
            get { return _canvas._horScrollbar.Visibility == Visibility.Visible; }
            set { _canvas._horScrollbar.Visibility = value ? Visibility.Visible : Visibility.Hidden; }
        }

        public bool ScrollBarVerticalVisible
        {
            get { return _canvas._verScrollbar.Visibility == Visibility.Visible; }
            set { _canvas._verScrollbar.Visibility = value ? Visibility.Visible : Visibility.Hidden; }
        }

        public double ScrollBarHorizontalMaximum
        {
            get { return _canvas._horScrollbar.Maximum; }
            set { _canvas._horScrollbar.Maximum = value; }
        }

        public double ScrollBarHorizontalMinimum
        {
            get { return _canvas._horScrollbar.Minimum; }
            set { _canvas._horScrollbar.Minimum = value; }
        }

        public double ScrollBarHorizontalValue
        {
            get { return _canvas._horScrollbar.Value; }
            set { _canvas._horScrollbar.Value = value; }
        }

        public double ScrollBarHorizontalLargeChange
        {
            get { return _canvas._horScrollbar.LargeChange; }
            set
            {
                _canvas._horScrollbar.LargeChange = value;
                _canvas._horScrollbar.ViewportSize = value;
            }
        }

        public double ScrollBarVerticalMaximum
        {
            get { return _canvas._verScrollbar.Maximum; }
            set { _canvas._verScrollbar.Maximum = value; }
        }

        public double ScrollBarVerticalMinimum
        {
            get { return _canvas._verScrollbar.Minimum; }
            set { _canvas._verScrollbar.Minimum = value; }
        }

        public double ScrollBarVerticalValue
        {
            get { return _canvas._verScrollbar.Value; }
            set { _canvas._verScrollbar.Value = value; }
        }

        public double ScrollBarVerticalLargeChange
        {
            get { return _canvas._verScrollbar.LargeChange; }
            set
            {
                _canvas._verScrollbar.LargeChange = value;
                _canvas._verScrollbar.ViewportSize = value;
            }
        }

        #endregion

        #region ITimerSupportedControlInterface Members

        public void StartTimer()
        {
            throw new NotImplementedException();
        }

        public void StopTimer()
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}