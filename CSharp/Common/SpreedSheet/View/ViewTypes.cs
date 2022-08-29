namespace SpreedSheet.View
{
    public enum ViewTypes
    {
        None = 0x0,
        Cells = 0x1,

        ColumnHeader = 0x2,
        RowHeader = 0x4,
        LeadHeader = ColumnHeader | RowHeader,

        ColOutline = 0x10,
        RowOutline = 0x20,
        Outlines = ColOutline | RowOutline
    }
}