using SpreedSheet.View;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace SpreedSheet.Rendering
{
    /// <summary>
    ///     Represents the platform no-associated drawing context.
    /// </summary>
    public abstract class DrawingContext
    {
        //internal DrawingContext(Worksheet worksheet, DrawMode drawMode)
        //	: this(worksheet, drawMode, null)
        //{
        //}

        internal DrawingContext(Worksheet worksheet, DrawMode drawMode, IRenderer r)
        {
            Worksheet = worksheet;
            DrawMode = drawMode;
            Graphics = r;
        }

        /// <summary>
        ///     Get current instance of worksheet.
        /// </summary>
        public Worksheet Worksheet { get; }

        /// <summary>
        ///     Platform independent drawing context.
        /// </summary>
        public IGraphics Graphics { get; internal set; }

        internal IRenderer Renderer
        {
            get { return (IRenderer)Graphics; }
        }

        internal IView CurrentView { get; set; }

        /// <summary>
        ///     Draw mode that decides what kind of content will be drawn during this drawing event.
        /// </summary>
        public DrawMode DrawMode { get; }
    }
}