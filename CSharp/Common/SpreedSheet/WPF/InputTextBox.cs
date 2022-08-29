using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SpreedSheet.Core.Enum;
using unvell.ReoGrid;

namespace SpreedSheet.WPF
{
    public class InputTextBox : TextBox
    {
        public SheetControl Owner { get; set; }
        internal bool TextWrap { get; set; }
        internal Size CellSize { get; set; }
        internal GridVerAlign VAlign { get; set; }

        protected override void OnLostFocus(RoutedEventArgs e)
        {
            var sheet = Owner.ActiveWorksheet;

            if (sheet.CurrentEditingCell != null && Visibility == Visibility.Visible)
            {
                sheet.EndEdit(Text);
                Visibility = Visibility.Hidden;
            }

            base.OnLostFocus(e);
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            var sheet = Owner.ActiveWorksheet;

            // in single line text
            if (!TextWrap && Text.IndexOf('\n') == -1)
            {
                Action moveAction = null;

                if (e.Key == Key.Up)
                    moveAction = () => sheet.MoveSelectionUp();
                else if (e.Key == Key.Down)
                    moveAction = () => sheet.MoveSelectionDown();
                else if (e.Key == Key.Left && SelectionStart == 0)
                    moveAction = () => sheet.MoveSelectionLeft();
                else if (e.Key == Key.Right && SelectionStart == Text.Length)
                    moveAction = () => sheet.MoveSelectionRight();
                if (moveAction != null)
                {
                    sheet.EndEdit(Text);
                    moveAction();
                    e.Handled = true;
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            var sheet = Owner.ActiveWorksheet;

            if (sheet.CurrentEditingCell != null && Visibility == Visibility.Visible)
            {
                if ((Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                    && e.Key == Key.Enter)
                {
                    var str = Text;
                    var selstart = SelectionStart;
                    str = str.Insert(SelectionStart, Environment.NewLine);
                    Text = str;
                    SelectionStart = selstart + Environment.NewLine.Length;
                }
                else if (!Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl) && e.Key == Key.Enter)
                {
                    sheet.EndEdit(Text);
                    sheet.MoveSelectionForward();
                    e.Handled = true;
                }
                else if (e.Key == Key.Enter)
                {
                    // TODO: auto adjust row height
                }
                // shift + tab
                else if ((Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) && e.Key == Key.Tab)
                {
                    sheet.EndEdit(Text);
                    sheet.MoveSelectionBackward();
                    e.Handled = true;
                }
                // tab
                else if (e.Key == Key.Tab)
                {
                    sheet.EndEdit(Text);
                    sheet.MoveSelectionForward();
                    e.Handled = true;
                }
                else if (e.Key == Key.Escape)
                {
                    sheet.EndEdit(EndEditReason.Cancel);
                    e.Handled = true;
                }
            }
        }

        protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            base.OnLostKeyboardFocus(e);
            Owner.ActiveWorksheet.EndEdit(Text, EndEditReason.NormalFinish);
        }

        protected override void OnTextChanged(TextChangedEventArgs e)
        {
            base.OnTextChanged(e);
            Text = Owner.ActiveWorksheet.RaiseCellEditTextChanging(Text);
        }

        protected override void OnPreviewTextInput(TextCompositionEventArgs e)
        {
            if (e.Text.Length > 0)
            {
                int inputChar = e.Text[0];
                if (inputChar != Owner.ActiveWorksheet.RaiseCellEditCharInputed(inputChar)) e.Handled = true;
            }

            base.OnPreviewTextInput(e);
        }
    }
}