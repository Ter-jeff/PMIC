using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl
{
    public abstract class MainBase : IDisposable
    {
        protected Dictionary<IgxlSheet, string> IgxlSheets = new Dictionary<IgxlSheet, string>();

        public void Dispose()
        {
            GC.Collect();
            for (var j = 0; j < GC.MaxGeneration; j++)
            {
                GC.Collect(j);
                GC.WaitForPendingFinalizers();
            }

            GC.SuppressFinalize(this);
        }
    }
}