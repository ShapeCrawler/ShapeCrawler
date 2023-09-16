using ShapeCrawler.Drawing;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     Represents interface of AutoShape.
/// </summary>
internal interface IRootSlideAutoShape : IShape
{     
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    void Duplicate();
}