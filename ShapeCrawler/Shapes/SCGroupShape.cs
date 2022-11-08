using OneOf;
using ShapeCrawler.Collections;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

internal class SCGroupShape : SlideShape, IGroupShape
{
    private readonly P.GroupShape pGroupShape;
    private readonly OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide;

    public SCGroupShape(P.GroupShape pGroupShape, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, Shape groupShape)
        : base(pGroupShape, oneOfSlide, groupShape)
    {
        this.pGroupShape = pGroupShape;
        this.oneOfSlide = oneOfSlide;
    }

    public IGroupedShapeCollection Shapes => GroupedShapeCollection.Create(this.pGroupShape, this.oneOfSlide, this);

    public override SCShapeType ShapeType => SCShapeType.GroupShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}