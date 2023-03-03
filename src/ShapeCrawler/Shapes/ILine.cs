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
public interface ILine : IShape
{
}

internal sealed class SCLine : SCShape, ILine
{
    public SCLine(
        OpenXmlCompositeElement childOfPShapeTree,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(childOfPShapeTree, parentSlideObject, parentShapeCollection)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.ConnectionShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}