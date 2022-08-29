using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Interaction
{
    public struct ResizeThumb
    {
        public ResizeThumbPosition Position;

        public Point Point;

        public ResizeThumb(ResizeThumbPosition position, Point point)
        {
            Position = position;
            Point = point;
        }

        public ResizeThumb(ResizeThumbPosition position, double width, double height)
        {
            Position = position;
            Point = new Point(width, height);
        }
    }
}