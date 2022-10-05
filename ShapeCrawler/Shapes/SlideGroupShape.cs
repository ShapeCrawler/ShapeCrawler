using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a group shape on a Slide.
/// </summary>
internal class SlideGroupShape : SlideShape, IGroupShape
{
    private readonly GroupShape pGroupShape;
    private readonly SCSlide slide;

    public SlideGroupShape(P.GroupShape pGroupShape, SCSlide slide, Shape groupShape)
        : base(pGroupShape, slide, groupShape)
    {
        this.pGroupShape = pGroupShape;
        this.slide = slide;
    }

    public IGroupedShapeCollection Shapes => GroupedShapeCollection.Create(this.pGroupShape, this.slide, this);

    public override SCShapeType ShapeType => SCShapeType.GroupShape;
}