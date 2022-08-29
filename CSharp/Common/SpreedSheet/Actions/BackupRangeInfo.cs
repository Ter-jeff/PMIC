#define WPF

namespace unvell.ReoGrid.Actions
{
    internal struct BackupRangeInfo
    {
        internal int start;
        internal int count;

        public BackupRangeInfo(int start, int count)
        {
            this.start = start;
            this.count = count;
        }
    }
}