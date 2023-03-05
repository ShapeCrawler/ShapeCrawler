using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using SkiaSharp;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine : IAutoShape
{
    /// <summary>
    ///    Gets the start point of the line.
    /// </summary>
    SCPoint StartPoint { get; }
    
    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    SCPoint EndPoint { get; }
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
    
    public SCPoint StartPoint => this.GetStartPoint();
    
    public SCPoint EndPoint => this.GetEndPoint();

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    private SCPoint GetStartPoint()
    {
        return new SCPoint(this.X, this.Y);
    }
    
    private SCPoint GetEndPoint()
    {
        var x = this.X + this.Width;
        var y = this.Y + this.Height;

        return new SCPoint(x, y);
    }
}