using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public abstract class IgxlRow
    {
        public bool IsBackup { get; set; }
        public string ColumnA { get; set; }
    }
}