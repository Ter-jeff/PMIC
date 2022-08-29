#define WPF

using System;
using System.Collections;
using System.Collections.Generic;
using SpreedSheet.Core;
using unvell.ReoGrid.Graphics;

namespace unvell.ReoGrid
{
    partial class Worksheet
    {
        private static readonly SolidColor[] NamedRangeHighlightColors =
        {
            SolidColor.Blue, SolidColor.Green, SolidColor.Purple,
            SolidColor.Brown, SolidColor.SeaGreen, SolidColor.Orange,
            SolidColor.IndianRed
        };

        internal HighlightRange focusHighlightRange;

        internal byte focusHighlightRangeRunningOffset = 0;

        private HighlightRangeCollection highlightRangeCollection;

        internal List<HighlightRange> highlightRanges = new List<HighlightRange>();

        private byte rangeHighlightColorCounter;

        /// <summary>
        ///     Get or set the focus highlight range
        /// </summary>
        public HighlightRange FocusHighlightRange
        {
            get { return focusHighlightRange; }
            set
            {
                focusHighlightRange = value;

                if (focusHighlightRange != null) controlAdapter.StartTimer();
            }
        }

        /// <summary>
        ///     Collection of highlighted ranges
        /// </summary>
        public ICollection<HighlightRange> HighlightRanges
        {
            get
            {
                if (highlightRangeCollection == null) highlightRangeCollection = new HighlightRangeCollection(this);
                return highlightRangeCollection;
            }
        }

        internal SolidColor GetNextAvailableHighlightRangeColor()
        {
            var color = NamedRangeHighlightColors[rangeHighlightColorCounter++];

            if (rangeHighlightColorCounter >= NamedRangeHighlightColors.Length) rangeHighlightColorCounter = 0;

            return color;
        }

        /// <summary>
        ///     Start pick and create a highlight range on spreadsheet.
        /// </summary>
        public void StartCreateHighlightRange()
        {
            operationStatus = OperationStatus.HighlightRangeCreate;
        }

        /// <summary>
        ///     Create highlight range from specified range position.
        /// </summary>
        /// <param name="address">Address or name to locate the range on worksheet.</param>
        /// <returns>Instance of highlight range created on the worksheet.</returns>
        public HighlightRange CreateHighlightRange(string addressOrName)
        {
            return CreateHighlightRange(addressOrName, GetNextAvailableHighlightRangeColor());
        }

        /// <summary>
        ///     Create highlight range at specified position
        /// </summary>
        /// <param name="addressOrName">Address or name to locate a range on worksheet</param>
        /// <param name="color">Color of the hihglight range displayed on worksheet</param>
        /// <returns>Instace of highlight range created in this worksheet</returns>
        public HighlightRange CreateHighlightRange(string addressOrName, SolidColor color)
        {
            if (RangePosition.IsValidAddress(addressOrName))
                return CreateHighlightRange(new RangePosition(addressOrName), color);

            NamedRange range;
            if (registeredNamedRanges.TryGetValue(addressOrName, out range))
                return CreateHighlightRange(range.Position, color);
            throw new InvalidAddressException(addressOrName);
        }

        /// <summary>
        ///     Create highlight range at specified position
        /// </summary>
        /// <param name="range">Range on worksheet to be highlight</param>
        /// <param name="color">Color of the hihglight range displayed on worksheet</param>
        /// <returns>Instace of highlight range created in this worksheet</returns>
        public HighlightRange CreateHighlightRange(RangePosition range, SolidColor color)
        {
            return new HighlightRange(this, range) { HighlightColor = color };
        }

        /// <summary>
        ///     Crearte and display a highlighted range at specified position on worksheet
        /// </summary>
        /// <param name="address">Address or name to locate a range on worksheet</param>
        /// <returns>Instance of highlight range on worksheet</returns>
        public HighlightRange AddHighlightRange(string address)
        {
            if (RangePosition.IsValidAddress(address)) return AddHighlightRange(new RangePosition(address));

            if (RGUtility.IsValidName(address))
            {
                NamedRange refRange;
                if (registeredNamedRanges.TryGetValue(address, out refRange)) return AddHighlightRange(refRange);
            }

            return null;
        }

        /// <summary>
        ///     Crearte and display a highlighted range at specified position on worksheet
        /// </summary>
        /// <param name="range">Position to add highlighted range</param>
        /// <returns>Instance of highlight range on worksheet</returns>
        public HighlightRange AddHighlightRange(RangePosition range)
        {
            for (var i = 0; i < highlightRanges.Count; i++)
            {
                var hr = highlightRanges[i];

                if (hr.StartPos == range.StartPos && hr.EndPos == range.EndPos) highlightRanges.RemoveAt(i);
            }

            var rrange = new HighlightRange(this, range);
            AddHighlightRange(rrange);

            return rrange;
        }

        internal void AddHighlightRange(HighlightRange range)
        {
            highlightRanges.Add(range);
            RequestInvalidate();
        }

        /// <summary>
        ///     Remove a highlighted range from specified address
        /// </summary>
        /// <param name="address">address to remove highlighted range</param>
        /// <returns>true if range removed successfully</returns>
        public bool RemoveHighlightRange(string address)
        {
            if (RangePosition.IsValidAddress(address)) return RemoveHighlightRange(new RangePosition(address));

            if (RGUtility.IsValidName(address))
            {
                NamedRange refRange;
                if (registeredNamedRanges.TryGetValue(address, out refRange)) return RemoveHighlightRange(refRange);
            }

            return false;
        }

        /// <summary>
        ///     Remove a highlighted range from specified position
        /// </summary>
        /// <param name="range">position to remove highlighted range</param>
        /// <returns>true if range removed successfully</returns>
        public bool RemoveHighlightRange(RangePosition range)
        {
            var found = false;

            for (var i = 0; i < highlightRanges.Count; i++)
                if (highlightRanges[i].StartPos == range.StartPos
                    && highlightRanges[i].EndPos == range.EndPos)
                {
                    highlightRanges.RemoveAt(i);
                    found = true;
                }

            if (found) RequestInvalidate();

            return found;
        }

        /// <summary>
        ///     Remove all highlighted ranges from current spreadsheet
        /// </summary>
        public void RemoveAllHighlightRanges()
        {
            if (highlightRanges != null
                && highlightRanges.Count > 0)
            {
                highlightRanges.Clear();
                RequestInvalidate();
            }
        }

        /// <summary>
        ///     Check whether a range specified by position is added into current spreadsheet
        /// </summary>
        /// <param name="range">range to be checked</param>
        /// <returns>true if specified range is alreay added</returns>
        public bool HasHighlightRange(RangePosition range)
        {
            for (var i = 0; i < highlightRanges.Count; i++)
                if (highlightRanges[i].StartPos == range.StartPos
                    && highlightRanges[i].EndPos == range.EndPos)
                {
                    highlightRanges.RemoveAt(i);
                    return true;
                }

            return false;
        }

        /// <summary>
        ///     Threading to update frames of focus highlighted range
        /// </summary>
        public void TimerRun()
        {
            if (focusHighlightRange == null)
                controlAdapter.StopTimer();
            else
                // todo: performance optimum (reduce repaint region by finding changed views)
                RequestInvalidate();
        }

        #region HighlightRangeCollection

        /// <summary>
        ///     Collection of highlighted range
        /// </summary>
        internal class HighlightRangeCollection : ICollection<HighlightRange>
        {
            public HighlightRangeCollection(Worksheet grid)
            {
                Worksheet = grid;
            }

            internal Worksheet Worksheet { get; }

            public void Add(HighlightRange item)
            {
                Worksheet.AddHighlightRange(item);
            }

            public void Clear()
            {
                Worksheet.RemoveAllHighlightRanges();
            }

            public bool Contains(HighlightRange item)
            {
                return Worksheet.HasHighlightRange(item);
            }

            public void CopyTo(HighlightRange[] array, int arrayIndex)
            {
                Worksheet.highlightRanges.CopyTo(array, arrayIndex);
            }

            public int Count
            {
                get { return Worksheet.highlightRanges.Count; }
            }

            public bool IsReadOnly
            {
                get { throw new NotImplementedException(); }
            }

            public bool Remove(HighlightRange item)
            {
                return Worksheet.RemoveHighlightRange(item);
            }

            public IEnumerator<HighlightRange> GetEnumerator()
            {
                return Worksheet.highlightRanges.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return Worksheet.highlightRanges.GetEnumerator();
            }

            public HighlightRange Add(string address)
            {
                return Worksheet.AddHighlightRange(address);
            }

            public HighlightRange Add(RangePosition range)
            {
                return Worksheet.AddHighlightRange(range);
            }
        }

        #endregion // HighlightRangeCollection
    }

    /// <summary>
    ///     Highlight range reference to spreadsheet
    /// </summary>
    public class HighlightRange : ReferenceRange
    {
        private SolidColor highlightColor;

        internal HighlightRange(Worksheet worksheet, CellPosition startPos, CellPosition endPos)
            : base(worksheet, startPos, endPos)
        {
            HighlightColor = worksheet.GetNextAvailableHighlightRangeColor();
        }

        internal HighlightRange(Worksheet worksheet, Cell startCell, Cell endCell)
            : this(worksheet, startCell.Position, endCell.Position)
        {
        }

        internal HighlightRange(Worksheet worksheet, string address)
            : this(worksheet, new RangePosition(address))
        {
        }

        internal HighlightRange(Worksheet worksheet, RangePosition range)
            : this(worksheet, range.StartPos, range.EndPos)
        {
        }

        /// <summary>
        ///     Highlight color to display range on spreadsheet
        /// </summary>
        public SolidColor HighlightColor
        {
            get { return highlightColor; }
            set
            {
                highlightColor = value;

                Worksheet.RequestInvalidate();
            }
        }

        /// <summary>
        ///     Decide whether this range is hover.
        /// </summary>
        public bool Hover { get; set; }

        internal int RunnerOffset { get; set; }
    }
}