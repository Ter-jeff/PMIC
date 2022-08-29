#define WPF

#if DRAWING
using System;

#if WINFORM
using RGFloat = System.Single;
#else
using RGFloat = System.Double;
#endif // WINFORM

using unvell.ReoGrid.Rendering;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid.Drawing.Shapes
{
#region Rectangle
	/// <summary>
	/// Represents regular rectangle drawing object.
	/// </summary>
	public class RectangleShape : ShapeObject
	{
	}
#endregion // Rectangle

#region Ellipse
	/// <summary>
	/// Represents ellipse shape object.
	/// </summary>
	public class EllipseShape : ShapeObject
	{
		/// <summary>
		/// Render ellipse shape to graphics context.
		/// </summary>
		/// <param name="dc">Platform no-associated drawing context.</param>
		protected override void OnPaint(DrawingContext dc)
		{
			var g = dc.Graphics;

			var clientRect = this.ClientBounds;

			if (clientRect.Width > 0 && clientRect.Height > 0)
			{
				if (!this.FillColor.IsTransparent)
				{
					g.FillEllipse(this.FillColor, clientRect);
				}

				if (!this.LineColor.IsTransparent)
				{
					g.DrawEllipse(this.LineColor, clientRect);
				}
			}
		}
	}
#endregion // Ellipse
}

#endif // DRAWING