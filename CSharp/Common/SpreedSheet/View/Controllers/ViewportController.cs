using System.Collections.Generic;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View.Controllers
{
    internal class ViewportController : IViewportController
    {
        #region Draw

        public virtual void Draw(CellDrawingContext dc)
        {
            if (view.Visible && view.Width > 0 && view.Height > 0) view.Draw(dc);

            Worksheet.viewDirty = false;
        }

        #endregion

        #region Constructor

        public Worksheet Worksheet { get; }

        public ViewportController(Worksheet sheet)
        {
            Worksheet = sheet;

            view = new View(this);
            view.Children = new List<IView>();
        }

        #endregion // Constructor

        #region Bounds

        //private RGRect bounds;
        public virtual Rectangle Bounds
        {
            get { return view.Bounds; }
            set { view.Bounds = value; }
        }

        //public virtual RGSize Size { get { return this.Bounds.Size; }  }
        //public virtual RGIntDouble Left { get { return this.Bounds.X; } }
        //public virtual RGIntDouble Top { get { return this.Bounds.Y; } }
        //public virtual RGIntDouble Width { get { return this.Bounds.Width; } }
        //public virtual RGIntDouble Height { get { return this.Bounds.Height; } }
        //public virtual RGIntDouble Right { get { return this.Bounds.Right; } }
        //public virtual RGIntDouble Bottom { get { return this.Bounds.Bottom; } }

        public virtual double ScaleFactor
        {
            get { return View == null ? 1f : View.ScaleFactor; }
            set
            {
                if (View != null) View.ScaleFactor = value;
            }
        }

        #endregion

        #region Viewport Management

        protected IView view;

        public virtual IView View
        {
            get { return view; }
            set { view = value; }
        }

        internal virtual void AddView(IView view)
        {
            this.view.Children.Add(view);
            view.ViewportController = this;
        }

        internal virtual void InsertView(IView before, IView viewport)
        {
            var views = view.Children;

            var index = views.IndexOf(before);
            if (index > 0 && index < views.Count)
                views.Insert(index, viewport);
            else
                views.Add(viewport);
        }

        internal virtual void InsertView(int index, IView viewport)
        {
            view.Children.Insert(index, viewport);
        }

        internal virtual bool RemoveView(IView view)
        {
            if (this.view.Children.Remove(view))
            {
                view.ViewportController = null;
                return true;
            }

            return false;
        }

        protected ViewTypes viewsVisible = ViewTypes.LeadHeader;

        internal bool IsViewVisible(ViewTypes head)
        {
            return (viewsVisible & head) == head;
        }

        public virtual void SetViewVisible(ViewTypes head, bool visible)
        {
            if (visible)
                viewsVisible |= head;
            else
                viewsVisible &= ~head;
        }

        #endregion

        #region Update

        public virtual void UpdateController()
        {
        }

        public virtual void Reset()
        {
        }

        public virtual void Invalidate()
        {
            if (Worksheet != null) Worksheet.RequestInvalidate();
        }

        #endregion

        #region Focus

        public virtual IView FocusView { get; set; }

        public virtual IUserVisual FocusVisual { get; set; }

        #endregion

        #region UI Handle

        public virtual bool OnMouseDown(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            if (!isProcessed)
            {
                var targetView = view.GetViewByPoint(location);

                if (targetView != null) isProcessed = targetView.OnMouseDown(targetView.PointToView(location), buttons);
            }

            return isProcessed;
        }

        public virtual bool OnMouseMove(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            if (FocusView != null) FocusView.OnMouseMove(FocusView.PointToView(location), buttons);

            if (!isProcessed)
            {
                var targetView = view.GetViewByPoint(location);

                if (targetView != null) isProcessed = targetView.OnMouseMove(targetView.PointToView(location), buttons);
            }

            return isProcessed;
        }

        public virtual bool OnMouseUp(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            if (FocusView != null) isProcessed = FocusView.OnMouseUp(FocusView.PointToView(location), buttons);

            return isProcessed;
        }

        public virtual bool OnMouseDoubleClick(Point location, MouseButtons buttons)
        {
            var isProcessed = false;

            var targetView = FocusView != null
                ? FocusView
                : view.GetViewByPoint(location);

            if (targetView != null)
                isProcessed = targetView.OnMouseDoubleClick(targetView.PointToView(location), buttons);

            return isProcessed;
        }

        public virtual bool OnKeyDown(KeyCode key)
        {
            return false;
        }

        public virtual void SetFocus()
        {
        }

        public virtual void FreeFocus()
        {
        }

        #endregion
    }
}