using System.Collections.Generic;
using SpreedSheet.Interaction;
using SpreedSheet.Rendering;
using SpreedSheet.View.Controllers;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.View
{
    internal class View : IView
    {
        private double _scaleFactor = 1f;

        public View()
        {
            Visible = true;
            PerformTransform = true;
        }

        public View(IViewportController vc)
            : this()
        {
            ViewportController = vc;
        }

        public IViewport Parent { get; set; }

        public IViewportController ViewportController { get; set; }

        public virtual double ScaleFactor
        {
            get { return _scaleFactor; }
            set { _scaleFactor = value; }
        }

        public virtual bool Visible { get; set; }

        public virtual bool PerformTransform { get; set; }

        public virtual void Draw(CellDrawingContext dc)
        {
            DrawChildren(dc);
        }

        public virtual void DrawChildren(CellDrawingContext dc)
        {
            if (Children != null)
                foreach (var view in Children)
                    if (view.Visible)
                    {
                        dc.CurrentView = view;
                        view.Draw(dc);
                        dc.CurrentView = null;
                    }
        }

        public virtual Point PointToView(Point p)
        {
            return new Point(
                (p.X - _bounds.Left) / _scaleFactor,
                (p.Y - _bounds.Top) / _scaleFactor);
        }

        public virtual Point PointToController(Point p)
        {
            return new Point(
                p.X * _scaleFactor + _bounds.Left,
                p.Y * _scaleFactor + _bounds.Top);
        }

        public virtual IView GetViewByPoint(Point p)
        {
            var child = GetChildrenByPoint(p);

            if (child != null)
                return child;
            return _bounds.Contains(p) ? this : null;
        }

        public virtual void Invalidate()
        {
            if (ViewportController != null) ViewportController.Invalidate();
        }

        public virtual bool OnMouseDown(Point location, MouseButtons buttons)
        {
            return false;
        }

        public virtual bool OnMouseMove(Point location, MouseButtons buttons)
        {
            return false;
        }

        public virtual bool OnMouseUp(Point location, MouseButtons buttons)
        {
            return false;
        }

        public virtual bool OnMouseDoubleClick(Point location, MouseButtons buttons)
        {
            return false;
        }

        public virtual bool OnKeyDown(KeyCode key)
        {
            return false;
        }

        public virtual void UpdateView()
        {
            if (Children != null)
                foreach (var child in Children)
                    child.UpdateView();
        }

        public virtual void SetFocus()
        {
            ViewportController.FocusView = this;
        }

        public virtual void FreeFocus()
        {
            ViewportController.FocusView = null;
        }

        public IList<IView> Children { get; set; }

        public virtual IView GetChildrenByPoint(Point p)
        {
            if (Children != null
                && Children.Count > 0)
                for (var i = Children.Count - 1; i >= 0; i--)
                {
                    var child = Children[i];
                    if (!child.Visible) continue;

                    var view = child.GetViewByPoint(p);

                    if (view != null) return view;
                }

            return null;
        }

        #region Bounds

        private Rectangle _bounds = new Rectangle(0, 0, 0, 0);

        public virtual Rectangle Bounds
        {
            get { return _bounds; }
            set { _bounds = value; }
        }

        public virtual double Top
        {
            get { return _bounds.Top; }
            set { _bounds.Y = value; }
        }

        public virtual double Left
        {
            get { return _bounds.Left; }
            set { _bounds.X = value; }
        }

        public virtual double Right
        {
            get { return _bounds.Right; }
        }

        public virtual double Bottom
        {
            get { return _bounds.Bottom; }
        }

        public virtual double Width
        {
            get { return _bounds.Width; }
            set { _bounds.Width = value; }
        }

        public virtual double Height
        {
            get { return _bounds.Height; }
            set { _bounds.Height = value; }
        }

        #endregion
    }
}