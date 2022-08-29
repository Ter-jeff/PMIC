#define WPF

using System;
using SpreedSheet.Core.Enum;

namespace unvell.ReoGrid.DataFormat
{
    internal class TextDataFormatter : IDataFormatter
    {
        public string FormatCell(Cell cell)
        {
            if (cell.InnerStyle.HAlign == GridHorAlign.General) cell.RenderHorAlign = GridRenderHorAlign.Left;

            return Convert.ToString(cell.InnerData);
        }

        public bool PerformTestFormat()
        {
            return false;
        }
    }
}