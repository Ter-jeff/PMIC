#define WPF

#if PRINT
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace unvell.ReoGrid.Print
{
	internal interface IPrintSession
#if WINFORM
		: IDisposable
#endif // WINFORM
	{
		IList<Worksheet> Worksheets { get; }

#if WINFORM
		System.Drawing.Printing.PrintDocument PrintDocument { get; }
#endif // WINFORM
	}
}

#endif // PRINT