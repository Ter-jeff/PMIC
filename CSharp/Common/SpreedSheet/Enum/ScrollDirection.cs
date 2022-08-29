namespace SpreedSheet.Enum
{
    internal enum ScrollDirection : byte
    {
        None = 0,
        Horizontal = 1,
        Vertical = 2,
        Both = Horizontal | Vertical
    }
}