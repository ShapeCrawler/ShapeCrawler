using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine : IAutoShape
{
}

internal sealed class SCLine : SCAutoShape, ILine
{
    public SCLine(
        TypedOpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideStructureOf,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollectionOf)
        : base(pShapeTreeChild, parentSlideStructureOf, parentShapeCollectionOf)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Line;

    public override ITextFrame? TextFrame => null;

    public override IShapeFill? Fill => null;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}