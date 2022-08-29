﻿#define WPF

#if PRINT
#if WPF

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using unvell.ReoGrid.Print;

namespace unvell.ReoGrid.Print
{
	partial class PrintSession
	{
		internal void Init() { }

		public void Dispose() { }

		/// <summary>
		/// Start output document to printer.
		/// </summary>
		public void Print()
		{
			throw new NotImplementedException("WPF Print is not implemented yet. Try use Windows Form version to print document as XPS file.");
		}
	}
}

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
	}
}

#endif // WPF

#endif // PRINT