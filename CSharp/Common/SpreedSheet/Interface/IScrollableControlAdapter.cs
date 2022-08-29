namespace SpreedSheet.Interface
{
    internal interface IScrollableControlAdapter
    {
        double ScrollBarHorizontalMaximum { get; set; }
        double ScrollBarHorizontalMinimum { get; set; }
        double ScrollBarHorizontalValue { get; set; }
        double ScrollBarHorizontalLargeChange { get; set; }
        double ScrollBarVerticalMaximum { get; set; }
        double ScrollBarVerticalMinimum { get; set; }
        double ScrollBarVerticalValue { get; set; }
        double ScrollBarVerticalLargeChange { get; set; }
    }
}