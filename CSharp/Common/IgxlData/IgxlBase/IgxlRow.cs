using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class IgxlRow
    {
        public string ColumnA { get; set; }
        public bool IsBackup { get; set; }
        public string SheetName { get; set; }
        public int RowNum { get; set; }
    }
}