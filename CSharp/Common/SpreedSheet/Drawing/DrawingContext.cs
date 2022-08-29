#define WPF

#if DRAWING
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using unvell.ReoGrid.Rendering;

namespace unvell.ReoGrid.Drawing
{
	/// <summary>
	/// Represents the floating objects drawing context.
	/// </summary>
	public class FloatingDrawingContext : unvell.ReoGrid.Rendering.DrawingContext
	{
		/// <summary>
		/// Get the current drawing object.
		/// </summary>
		public DrawingObject CurrentObject { get; private set; }

		internal FloatingDrawingContext(Worksheet worksheet, DrawMode drawMode, IRenderer r)
			: base(worksheet, drawMode, r)
		{
		}

		internal void SetCurrentObject(DrawingObject currentObject)
		{
			this.CurrentObject = currentObject;
		}
	}
}

#endif // DRAWING