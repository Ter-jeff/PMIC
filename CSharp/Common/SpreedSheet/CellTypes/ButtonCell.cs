using System;
using SpreedSheet.Core.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Representation for button of cell body.
    /// </summary>
    [Serializable]
    public class ButtonCell : CellBody
    {
        private string _defaultText;

        /// <summary>
        ///     Create button cell
        /// </summary>
        public ButtonCell()
        {
        }

        /// <summary>
        ///     Create button cell with specified text.
        /// </summary>
        /// <param name="defaultText"></param>
        public ButtonCell(string defaultText)
        {
            _defaultText = defaultText;
        }

        /// <summary>
        ///     Check whether or not button is pressed.
        /// </summary>
        public bool IsPressed { get; set; }

        /// <summary>
        ///     On body is setup to cell.
        /// </summary>
        /// <param name="cell">cell instance</param>
        public override void OnSetup(Cell cell)
        {
            // set text alignment to center
            if (cell != null)
            {
                if (cell.Worksheet != null)
                    cell.Worksheet.SetRangeStyles(cell.PositionAsRange,
                        new WorksheetRangeStyle
                        {
                            Flag = PlainStyleFlag.HorizontalAlign | PlainStyleFlag.VerticalAlign,
                            HAlign = GridHorAlign.Center,
                            VAlign = GridVerAlign.Middle
                        });

                // set default cell value
                if (!string.IsNullOrEmpty(_defaultText)) cell.Data = _defaultText;
            }
        }

        /// <summary>
        ///     Return false to disable edit operation for this cell.
        /// </summary>
        /// <returns>False to disable edit operation for this cell.</returns>
        public override bool OnStartEdit()
        {
            return false;
        }

        /// <summary>
        ///     Perform click operation.
        /// </summary>
        public virtual void PerformClick()
        {
            Click?.Invoke(this, null);
        }

        /// <summary>
        ///     Click event raised when user clicked on the button.
        /// </summary>
        public event EventHandler Click;

        /// <summary>
        ///     Clone a button cell from this object.
        /// </summary>
        /// <returns>New instance of button cell.</returns>
        public override ICellBody Clone()
        {
            return new ButtonCell(Cell.DisplayText);
        }

        #region Draw

        /// <summary>
        ///     Paint this cell body.
        /// </summary>
        /// <param name="dc">ReoGrid common drawing context</param>
        public override void OnPaint(CellDrawingContext dc)
        {
            if (Cell != null) DrawButton(dc);
            // get style
            //var style = this.Cell.InnerStyle;
            //var textColor = style.TextColor;
            // call core text drawing method
            dc.DrawCellText();
        }

        /// <summary>
        ///     Draw button surface.
        /// </summary>
        /// <param name="dc">Platform independence drawing context.</param>
        /// <param name="state">Button state.</param>
        protected virtual void DrawButton(CellDrawingContext dc)
        {
#if WINFORM
			ControlPaint.DrawButton(dc.Graphics.PlatformGraphics, (System.Drawing.Rectangle)Bounds,
				this.IsPressed ? ButtonState.Pushed :
				(this.Cell.IsReadOnly ? ButtonState.Inactive : ButtonState.Normal));
#elif WPF

            var g = dc.Graphics;

            //g.TranslateTransform(20f, 00f);

            var r = Bounds;
            g.DrawRectangle(r, SolidColor.Dark(StaticResources.SystemColor_ControlDark));

            //var r2 = new Rectangle(r.X, r.Y, r.Width - 1, r.Height - 1);
            var r3 = new Rectangle(r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2);
            g.FillRectangle(r3, StaticResources.SystemColor_Control);

            if (IsPressed)
            {
                //	g.DrawRectangle(r, StaticResources.SystemColor_ControlDark);

                //	r.X++; r.Y++; r.Width--; r.Height--;
                //	g.DrawRectangle(r, SolidColor.Dark(StaticResources.SystemColor_ControlDark));
                var r2 = new Rectangle(r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2);
                g.DrawRectangle(r2, SolidColor.Dark(StaticResources.SystemColor_ControlDark));
            }


            //var r2 = new Rectangle(r.X + 1, r.Y + 1, r.Width - 1, r.Height - 1);
            //g.DrawRectangle(r2, StaticResources.SystemColor_ControlDark);
#endif // WPF
        }

        #endregion Draw

        #region Mouse

        /// <summary>
        ///     Invoked when mouse down inside this body
        /// </summary>
        /// <param name="e">mouse event argument</param>
        /// <returns>true if event has been handled</returns>
        public override bool OnMouseDown(CellMouseEventArgs e)
        {
            if (Bounds.Contains(e.RelativePosition))
            {
                IsPressed = true;
                return true;
            }

            return false;
        }

        /// <summary>
        ///     Determine whether the button is pressed or released
        /// </summary>
        /// <param name="e">mouse event argument</param>
        /// <returns>true to notify spreadsheet that event has been handled</returns>
        public override bool OnMouseUp(CellMouseEventArgs e)
        {
            if (IsPressed)
            {
                if (Bounds.Contains(e.RelativePosition)) PerformClick();

                IsPressed = false;
                return true;
            }

            return false;
        }

        #endregion // Mouse

        #region Keyboard

        /// <summary>
        ///     Check when user pressed Space key to press button.
        /// </summary>
        /// <param name="keyCode">Virtual keys code that is converted from system platform.</param>
        /// <returns>True to notify spreadsheet that event has been handled</returns>
        public override bool OnKeyDown(KeyCode keyCode)
        {
            if (keyCode == KeyCode.Space)
            {
                IsPressed = true;
                return true;
            }

            return false;
        }

        /// <summary>
        ///     Check when user released Space key to restore button.
        /// </summary>
        /// <param name="keyCode">Virtual keys code that is converted from system platform.</param>
        /// <returns>True to notify spreadsheet that event has been handled.</returns>
        public override bool OnKeyUp(KeyCode keyCode)
        {
            if (keyCode == KeyCode.Space && IsPressed)
            {
                IsPressed = false;
                PerformClick();
                return true;
            }

            return false;
        }

        #endregion // Keyboard
    }
}