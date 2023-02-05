using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

/// <summary>
///     Represents a Rounded Rectangle shape.
/// </summary>
public interface IRoundedRectangle : IAutoShape
{
}

internal sealed class SCRoundedRectangle : SCAutoShape, IRoundedRectangle
{
    internal SCRoundedRectangle(
        P.Shape pShape, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection) 
        : base(pShape, parentSlideObject, parentShapeCollection)
    {
    }
}