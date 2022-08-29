using System;
using SpreedSheet.Core.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Representation for check box of cell body
    /// </summary>
    [Serializable]
    public class CheckBoxCell : ContentCellBody
    {
        #region Paint

        /// <summary>
        ///     Paint content of cell body.
        /// </summary>
        /// <param name="dc">Platform independency graphics context.</param>
        protected override void OnContentPaint(CellDrawingContext dc)
        {
#if WINFORM
			System.Windows.Forms.ButtonState bs = ButtonState.Normal;
			if (IsPressed) bs |= ButtonState.Pushed;
			if (IsChecked) bs |= ButtonState.Checked;

			ControlPaint.DrawCheckBox(dc.Graphics.PlatformGraphics, (System.Drawing.Rectangle)ContentBounds, bs);
#elif WPF
            var g = dc.Graphics;

            if (IsPressed) g.FillRectangle(ContentBounds, StaticResources.SystemColor_Control);

            g.DrawRectangle(ContentBounds, StaticResources.SystemColor_ControlDark);

            if (isChecked)
            {
                var x = ContentBounds.X;
                var y = ContentBounds.Y;
                var w = ContentBounds.Width;
                var h = ContentBounds.Height;

                var path = new PathGeometry();
                var figure = new PathFigure(new Point(x + w * 0.167, y + h * 0.546),
                    new[]
                    {
                        new LineSegment(new Point(x + w * 0.444, y + h * 0.712), false),
                        new LineSegment(new Point(x + w * 0.833, y + h * 0.157), false),
                        new LineSegment(new Point(x + w * 0.944, y + h * 0.323), false),
                        new LineSegment(new Point(x + w * 0.500, y + h * 0.934), false),
                        new LineSegment(new Point(x + w * 0.080, y + h * 0.712), false)
                    }, true);

                path.Figures.Add(figure);

                g.FillPath(StaticResources.SystemColor_WindowText, path);
            }
#endif // WPF
        }

        #endregion // Paint

        public override bool OnStartEdit()
        {
            return false;
        }

        /// <summary>
        ///     Handle event when data set into the cell of this body.
        /// </summary>
        /// <param name="data">Data inputted by user.</param>
        /// <returns>Data to be set into the cell.</returns>
        public override object OnSetData(object data)
        {
            IsChecked = data is bool && (bool)data;

            return base.OnSetData(data);
        }

        public override ICellBody Clone()
        {
            return new CheckBoxCell();
        }

        #region Constructor

        private bool initChecked;

        /// <summary>
        ///     Create check box cell body.
        /// </summary>
        public CheckBoxCell()
            : this(false)
        {
        }

        /// <summary>
        ///     Create check box cell body.
        /// </summary>
        /// <param name="initChecked">
        ///     Set the initial status. If this value is true, checkbox keep checked status when added into a
        ///     cell.
        /// </param>
        public CheckBoxCell(bool initChecked)
        {
            this.initChecked = initChecked;
        }

        /// <summary>
        ///     Override OnSetup method to set initial checked status.
        /// </summary>
        /// <param name="cell">The cell this body will set into.</param>
        public override void OnSetup(Cell cell)
        {
            if (cell != null)
            {
                if (initChecked) cell.Data = true;

                cell.Style.HAlign = GridHorAlign.Center;
                cell.Style.VAlign = GridVerAlign.Middle;
            }
        }

        #endregion // Constructor

        #region Mouse

        /// <summary>
        ///     Determines whether or not mouse or key pressed inside check box.
        /// </summary>
        protected virtual bool IsPressed { get; set; }

        /// <summary>
        ///     Handle the mouse down event.
        /// </summary>
        /// <param name="e">Arguments of mouse down event.</param>
        /// <returns>True if event has been handled; Otherwise return false.</returns>
        public override bool OnMouseDown(CellMouseEventArgs e)
        {
            if (ContentBounds.Contains(e.RelativePosition))
            {
                IsPressed = true;
                return true;
            }

            return false;
        }

        /// <summary>
        ///     Handle the mouse up event.
        /// </summary>
        /// <param name="e">Arguments of mouse up event.</param>
        /// <returns>True if event has been handled; Otherwise return false.</returns>
        public override bool OnMouseUp(CellMouseEventArgs e)
        {
            if (IsPressed)
            {
                IsPressed = false;

                if (ContentBounds.Contains(e.RelativePosition))
                {
                    ToggleCheckStatus();

                    RaiseClickEvent();
                }

                return true;
            }

            return false;
        }

        #endregion // Mouse

        #region Event

        /// <summary>
        ///     Event raied when user clicked inside check box.
        /// </summary>
        public event EventHandler Click;

        /// <summary>
        ///     Event raised when check status changed.
        /// </summary>
        public event EventHandler CheckChanged;

        /// <summary>
        ///     Raise the click event.
        /// </summary>
        protected virtual void RaiseClickEvent()
        {
            if (Click != null) Click(this, null);
        }

        /// <summary>
        ///     Raise the click event.
        /// </summary>
        protected virtual void RaiseCheckChangedEvent()
        {
            CheckChanged?.Invoke(this, null);
        }

        #endregion // Event

        #region Check Status

        /// <summary>
        ///     Toggle the check status of check box.
        /// </summary>
        public virtual void ToggleCheckStatus()
        {
            if (Cell != null
                && DisableWhenCellReadonly
                && Cell.IsReadOnly)
                return;

            IsChecked = !IsChecked;
        }

        /// <summary>
        ///     Check status.
        /// </summary>
        protected bool isChecked;

        /// <summary>
        ///     Get or set the check-status of check box.
        /// </summary>
        public virtual bool IsChecked
        {
            get { return isChecked; }
            set
            {
                if (isChecked != value)
                {
                    isChecked = value;

                    if (Cell != null && (Cell.InnerData as bool? ?? false) != value) Cell.Data = value;

                    CheckChanged?.Invoke(this, null);
                }
            }
        }

        #endregion // Check Status

        #region Keyboard

        /// <summary>
        ///     Handle keyboard down event.
        /// </summary>
        /// <param name="keyCode">Virtual keys that are pressed.</param>
        /// <returns>True if event has been marked as handled.</returns>
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
        ///     Handle keyboard up event.
        /// </summary>
        /// <param name="keyCode">Virtual keys code that is converted from system platform.</param>
        /// <returns>True if event has been handled; Otherwise return false.</returns>
        public override bool OnKeyUp(KeyCode keyCode)
        {
            if (IsPressed)
            {
                IsPressed = false;

                if (keyCode == KeyCode.Space) ToggleCheckStatus();

                return true;
            }

            return false;
        }

        #endregion // Keyboard
    }
}