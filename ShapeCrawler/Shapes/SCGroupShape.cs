using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories;
using ShapeCrawler.SlideMasters;
using OneOf;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a group shape on a Slide.
/// </summary>
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