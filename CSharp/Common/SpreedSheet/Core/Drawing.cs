#define WPF

#if DRAWING
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace unvell.ReoGrid
{
	using unvell.ReoGrid.Drawing;

	partial class Worksheet
	{
		internal WorksheetDrawingCanvas drawingCanvas;

		/// <summary>
		/// Access the collection of floating objects from worksheet.
		/// </summary>
		public IFloatingObjectCollection<IDrawingObject> FloatingObjects
		{
			get { return this.drawingCanvas.Children; }
		}
	}
}

#endif // DRAWING