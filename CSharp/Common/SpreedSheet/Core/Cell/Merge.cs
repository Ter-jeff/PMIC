#define WPF

using System.Diagnostics;
using SpreedSheet.Core;

namespace unvell.ReoGrid
{
    partial class Cell
    {
        private CellPosition mergeEndPos = CellPosition.Empty;
        private CellPosition mergeStartPos = CellPosition.Empty;

        internal CellPosition MergeStartPos
        {
            get { return mergeStartPos; }
            set
            {
#if DEBUG
                Debug.Assert(value.Row >= -1 && value.Col >= -1);
#endif

                mergeStartPos = value;
            }
        }

        internal CellPosition MergeEndPos
        {
            get { return mergeEndPos; }
            set
            {
#if DEBUG
                if ((value.Row > -1 && value.Col <= -1)
                    || (value.Row <= -1 && value.Col > -1))
                    Debug.Assert(false);

                Debug.Assert(value.Row >= -1 && value.Col >= -1);
#endif
                mergeEndPos = value;
            }
        }

        internal bool IsStartMergedCell
        {
            get { return InternalPos.Equals(MergeStartPos); }
        }

        internal bool IsEndMergedCell
        {
            get { return InternalPos.Row == mergeEndPos.Row && InternalPos.Col == mergeEndPos.Col; }
        }

        /// <summary>
        ///     Check whether this cell is merged cell
        /// </summary>
        public bool IsMergedCell
        {
            get { return IsStartMergedCell; }
        }

        /// <summary>
        ///     Check whether or not this cell is an valid cell, only valid cells can be set data and styles.
        ///     Cells merged by another cell will become invalid.
        /// </summary>
        public bool IsValidCell
        {
            get { return rowspan >= 1 && colspan >= 1; }
        }

        /// <summary>
        ///     Check whether or not this cell is inside a merged range
        /// </summary>
        public bool InsideMergedRange
        {
            get { return IsStartMergedCell || (rowspan == 0 && colspan == 0); }
        }
    }
}