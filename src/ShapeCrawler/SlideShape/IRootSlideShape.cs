using ShapeCrawler.Drawing;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     Represents interface of AutoShape.
/// </summary>
internal interface IRootSlideShape : IShape, ICopyableShape
{     
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    void Duplicate();
}