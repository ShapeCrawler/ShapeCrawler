using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

/// <summary>
///     Represents a Rectangle shape.
/// </summary>
public interface IRectangle : IAutoShape
{
}

internal sealed class SCRectangle : SCAutoShape, IRectangle
{
    internal SCRectangle(
        P.Shape pShape,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pShape, parentSlideObject, parentShapeCollection)
    {
    }
}