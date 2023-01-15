using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

/// <summary>
///     Represents a Rectangle shape.
/// </summary>
public interface IRectangle : IAutoShape
{
}

internal sealed class SCRectangle : AutoSCShape, IRectangle
{
    public SCRectangle(AutoShapeCollection autoShapeCollection, P.Shape pShape, SCGroupSCShape? groupShape) 
        : base(autoShapeCollection.ParentShapeCollection.ParentSlideObject, pShape, groupShape)
    {
    }
}