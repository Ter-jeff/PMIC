using System.Windows.Input;
using SpreedSheet.Interaction;

namespace SpreedSheet.WPF
{
    internal class WpfUtility
    {
        public static MouseButtons ConvertToUiMouseButtons(MouseEventArgs e)
        {
            var btn = MouseButtons.None;
            if (e.LeftButton == MouseButtonState.Pressed) btn |= MouseButtons.Left;
            if (e.MiddleButton == MouseButtonState.Pressed) btn |= MouseButtons.Middle;
            if (e.RightButton == MouseButtonState.Pressed) btn |= MouseButtons.Right;
            return btn;
        }
    }
}