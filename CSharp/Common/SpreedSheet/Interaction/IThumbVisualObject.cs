using System.Collections.Generic;

namespace SpreedSheet.Interaction
{
    internal interface IThumbVisualObject
    {
        IEnumerable<ResizeThumb> ThumbPoints { get; }
    }
}