using System;
using System.Collections.Generic;
using SpreedSheet.Rendering;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.Rendering;

namespace SpreedSheet.CellTypes
{
    #region Radio Group

    /// <summary>
    ///     Radio button group for toggling radios inside one group.
    /// </summary>
    [Serializable]
    public class RadioButtonGroup
    {
        public List<RadioButtonCell> RadioButtons { get; } = new List<RadioButtonCell>();

        /// <summary>
        ///     Add radio button into this group.
        /// </summary>
        /// <param name="cell"></param>
        public virtual void AddRadioButton(RadioButtonCell cell)
        {
            if (cell == null) return;

            if (!RadioButtons.Contains(cell)) RadioButtons.Add(cell);

            if (cell.RadioGroup != this) cell.RadioGroup = this;
        }

        /// <summary>
        ///     Check whether specified radio is contained by this group.
        /// </summary>
        /// <param name="cell">radio cell body to be checked.</param>
        /// <returns>true if the radio cell body is contained by this group.</returns>
        public virtual bool Contains(RadioButtonCell cell)
        {
            return RadioButtons.Contains(cell);
        }
    }

    #endregion // Radio Group

    /// <summary>
    ///     Representation for a radio button of cell body.
    /// </summary>
    [Serializable]
    public class RadioButtonCell : CheckBoxCell
    {
        /// <summary>
        ///     Create instance of radio button cell.
        /// </summary>
        public RadioButtonCell()
        {
        }

        /// <summary>
        ///     Get or set check status for radio button
        /// </summary>
        public override bool IsChecked
        {
            get { return isChecked; }
            set
            {
                if (isChecked != value)
                {
                    isChecked = value;

                    // uncheck other radios in same group
                    if (isChecked && radioGroup != null)
                        foreach (var other in radioGroup.RadioButtons)
                            if (other != this)
                                other.IsChecked = false;

                    if (Cell != null && (Cell.InnerData as bool? ?? false) != value) Cell.Data = value;

                    RaiseCheckChangedEvent();
                }
            }
        }

        /// <summary>
        ///     Toggle check status of radio-button. (Only work when radio button not be added into any groups)
        /// </summary>
        public override void ToggleCheckStatus()
        {
            if (!isChecked || radioGroup == null) base.ToggleCheckStatus();
        }

        /// <summary>
        ///     Paint content of cell body.
        /// </summary>
        /// <param name="dc">Platform independency graphics context.</param>
        protected override void OnContentPaint(CellDrawingContext dc)
        {
#if WINFORM
			System.Windows.Forms.ButtonState state = ButtonState.Normal;

			if (this.IsPressed) state |= ButtonState.Pushed;
			if (this.IsChecked) state |= ButtonState.Checked;

			ControlPaint.DrawRadioButton(dc.Graphics.PlatformGraphics,
				(System.Drawing.Rectangle)this.ContentBounds, state);

#elif WPF
            var g = dc.Graphics;

            var ox = ContentBounds.OriginX;
            var oy = ContentBounds.OriginY;

            var hw = ContentBounds.Width / 2;
            var hh = ContentBounds.Height / 2;
            var r = new Rectangle(ox - hw / 2, oy - hh / 2, hw, hh);
            g.DrawEllipse(StaticResources.SystemColor_ControlDark, r);

            if (IsPressed) g.FillEllipse(StaticResources.SystemColor_Control, r);

            if (isChecked)
            {
                var hhw = ContentBounds.Width / 4;
                var hhh = ContentBounds.Height / 4;
                r = new Rectangle(ox - hhw / 2, oy - hhh / 2, hhw, hhh);
                g.FillEllipse(StaticResources.SystemColor_WindowText, r);
            }
#endif // WINFORM
        }

        /// <summary>
        ///     Clone radio button from this object.
        /// </summary>
        /// <returns>New instance of radio button.</returns>
        public override ICellBody Clone()
        {
            return new RadioButtonCell();
        }

        #region Group

        private RadioButtonGroup radioGroup;

        /// <summary>
        ///     Radio groups for toggling other radios inside same group.
        /// </summary>
        public virtual RadioButtonGroup RadioGroup
        {
            get { return radioGroup; }
            set
            {
                if (value == null)
                {
                    RadioGroup = null;
                }
                else
                {
                    if (!value.Contains(this)) value.AddRadioButton(this);

                    radioGroup = value;
                }
            }
        }

        #endregion // Group
    }
}