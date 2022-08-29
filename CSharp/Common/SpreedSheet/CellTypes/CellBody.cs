using System;
using SpreedSheet.Core.Enum;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.CellTypes
{
    /// <summary>
    ///     Cell body
    /// </summary>
    [Serializable]
    public class CellBody : ICellBody
    {
        internal Cell InnerCell
        {
            set { Cell = value; }
        }

        /// <summary>
        ///     Owner cell contains this body.
        /// </summary>
        public Cell Cell { get; private set; }

        /// <summary>
        ///     When the body set into a cell.
        /// </summary>
        /// <param name="cell">Current owner cell</param>
        public virtual void OnSetup(Cell cell)
        {
            Cell = cell;
        }

        /// <summary>
        ///     Get cell body bounds rectangle.
        /// </summary>
        public virtual Rectangle Bounds { get; set; }

        /// <summary>
        ///     Determines whether or not become disable when owner cell is set as read-only. (Default is True)
        /// </summary>
        public virtual bool DisableWhenCellReadonly
        {
            get { return true; }
        }

        /// <summary>
        ///     Invoked when body boundary has been changed.
        /// </summary>
        public virtual void OnBoundsChanged()
        {
        }

        /// <summary>
        ///     Determines whether or not to allow capture the mouse when mouse down inside this body.
        /// </summary>
        /// <returns>True to allow caption; False to abort capture.</returns>
        public virtual bool AutoCaptureMouse()
        {
            return true;
        }

        /// <summary>
        ///     This method will be invoked when mouse button pressed inside the body bounds.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>Return true if event has been handled; Otherwise return false to recall default operations.</returns>
        public virtual bool OnMouseDown(CellMouseEventArgs e)
        {
            return false;
        }

        /// <summary>
        ///     This method will be invoked when mouse has been moved inside the body bounds.
        /// </summary>
        /// <param name="e">Mouse event argument</param>
        /// <returns>Return true if event has been handled; Otherwise false to recall default operations.</returns>
        public virtual bool OnMouseMove(CellMouseEventArgs e)
        {
            return false;
        }

        /// <summary>
        ///     This method will be invoked when any key released on this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>Return true if event has been handled; Otherwise return false to recall default operations.</returns>
        public virtual bool OnMouseUp(CellMouseEventArgs e)
        {
            return false;
        }

        /// <summary>
        ///     Invoked when mouse moved enter this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>True if event has been handled inside this body, otherwise false to recall built-in operations.</returns>
        public virtual bool OnMouseEnter(CellMouseEventArgs e)
        {
            return false;
        }

        /// <summary>
        ///     Invoked when mouse moved out from this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>True if event has been handled inside this body, otherwise false to recall built-in operations.</returns>
        public virtual bool OnMouseLeave(CellMouseEventArgs e)
        {
            return false;
        }

        /// <summary>
        ///     Invoked when mouse scrolled inside this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>True if event has been handled inside this body, otherwise false to recall built-in operations.</returns>
        public virtual void OnMouseWheel(CellMouseEventArgs e)
        {
        }

        /// <summary>
        ///     Invoked when any key pressed on this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>True if event has been handled inside this body, otherwise false to recall built-in operations.</returns>
        public virtual bool OnKeyDown(KeyCode e)
        {
            return false;
        }

        /// <summary>
        ///     Invoked when any key released on this body.
        /// </summary>
        /// <param name="e">Mouse event argument.</param>
        /// <returns>True if event has been handled inside this body, otherwise false to recall built-in operations.</returns>
        public virtual bool OnKeyUp(KeyCode e)
        {
            return false;
        }

        /// <summary>
        ///     Paint the content of body.
        /// </summary>
        /// <param name="dc">Platform independency graphics context.</param>
        public virtual void OnPaint(CellDrawingContext dc)
        {
            dc.DrawCellBackground();
            dc.DrawCellText();
        }

        /// <summary>
        ///     Invoked when cell of this body begin to edit. (Enter edit mode)
        /// </summary>
        /// <returns>True to allow edit; Otherwise false to cancel edit.</returns>
        public virtual bool OnStartEdit()
        {
            return true;
        }

        /// <summary>
        ///     Invoked when cell of this body finish edit.
        ///     Return data to be set into spreadsheet instead of user inputted.
        /// </summary>
        /// <param name="data">user inputted data.</param>
        /// <returns>new data to be into spreadsheet.</returns>
        public virtual object OnEndEdit(object data)
        {
            return data;
        }

        /// <summary>
        ///     Invoked when cell get focus.
        /// </summary>
        public virtual void OnGotFocus()
        {
        }

        /// <summary>
        ///     Invoked when cell lost focus.
        /// </summary>
        public virtual void OnLostFocus()
        {
        }

        /// <summary>
        ///     Invoked when cell data updating.
        /// </summary>
        /// <param name="data">Data to be updated.</param>
        /// <returns>New data that is used to replace the data inputted.</returns>
        public virtual object OnSetData(object data)
        {
            return data;
        }

        /// <summary>
        ///     Clone a cell body from this object.
        /// </summary>
        /// <returns>New instance of cell body.</returns>
        public virtual ICellBody Clone()
        {
            return new CellBody();
        }
    }

    /// <summary>
    ///     Represents a cell body that maintains a visual content region for child objects, such as Check-box or Radio-button.
    /// </summary>
    public abstract class ContentCellBody : CellBody
    {
        /// <summary>
        ///     Get or set child content bounds rectangle.
        /// </summary>
        public virtual Rectangle ContentBounds { get; set; }

        /// <summary>
        ///     Determines the preferred body size.
        /// </summary>
        protected virtual Size GetContentSize()
        {
            return new Size(17, 17);
        }

        /// <summary>
        ///     Handles when bounds changed.
        /// </summary>
        public override void OnBoundsChanged()
        {
            base.OnBoundsChanged();

            var contentRect = new Rectangle(new Point(0, 0), GetContentSize());

            if (Cell != null)
            {
                double x = 0, y = 0;

                switch (Cell.InnerStyle.HAlign)
                {
                    case GridHorAlign.Left:
                        x = Bounds.X + 1;
                        break;

                    case GridHorAlign.Center:
                        x = Bounds.X + (Bounds.Width - contentRect.Width) / 2;
                        break;

                    case GridHorAlign.Right:
                        x = Bounds.Right - contentRect.Width - 1;
                        break;
                }

                switch (Cell.InnerStyle.VAlign)
                {
                    case GridVerAlign.Top:
                        y = Bounds.Y + 1;
                        break;

                    case GridVerAlign.Middle:
                        y = Bounds.Y + (Bounds.Height - contentRect.Height) / 2;
                        break;

                    case GridVerAlign.Bottom:
                        y = Bounds.Bottom - contentRect.Height - 1;
                        break;
                }

                contentRect = new Rectangle(x, y, contentRect.Width, contentRect.Height);
            }
            else
            {
                contentRect = new Rectangle(Bounds.X + (Bounds.Width - contentRect.Width) / 2,
                    Bounds.Y + (Bounds.Height - contentRect.Height) / 2, contentRect.Width, contentRect.Height);
            }

            ContentBounds = contentRect;
        }

        /// <summary>
        ///     Paint cell body.
        /// </summary>
        /// <param name="dc">Platform independency graphics context.</param>
        public override void OnPaint(CellDrawingContext dc)
        {
            dc.DrawCellBackground();

            if (ContentBounds.Width > 0 || ContentBounds.Height > 0) OnContentPaint(dc);
        }

        /// <summary>
        ///     Paint content of cell body.
        /// </summary>
        /// <param name="dc">Platform independency graphics context.</param>
        protected virtual void OnContentPaint(CellDrawingContext dc)
        {
        }
    }
}