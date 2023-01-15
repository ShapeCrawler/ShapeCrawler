using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a connection shape.
/// </summary>
public interface IConnectionShape : IShape
{
}

internal sealed class SCConnectionSCShape : SCSlideShape, IConnectionShape
{
    public SCConnectionSCShape(OpenXmlCompositeElement childOfPShapeTree, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
        : base(childOfPShapeTree, oneOfSlide, null)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.ConnectionShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}