using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

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
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf)
        : base(pShapeTreeChild, slideOf, shapeCollectionOf)
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
        var horizontalFlip = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !.Transform2D!.HorizontalFlip?.Value;
        var flipH = horizontalFlip != null && horizontalFlip.Value;
        var verticalFlip = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !.Transform2D!.VerticalFlip?.Value;
        var flipV = verticalFlip != null && verticalFlip.Value;

        if (flipH && (this.Height == 0 || flipV))
        {
            return new SCPoint(this.X, this.Y);
        }
        
        if (flipH)
        {
            return new SCPoint(this.X + this.Width, this.Y);
        }
        
        return new SCPoint(this.X, this.Y);
    }
    
    private SCPoint GetEndPoint()
    {
        var horizontalFlip = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !.Transform2D!.HorizontalFlip?.Value;
        var flipH = horizontalFlip != null && horizontalFlip.Value;
        var verticalFlip = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !.Transform2D!.VerticalFlip?.Value;
        var flipV = verticalFlip != null && verticalFlip.Value;

        if(this.Width == 0)
        {
            return new SCPoint(this.X, this.Height);
        }
        
        if (flipH && this.Height == 0)
        {
            return new SCPoint(this.X - this.Width, this.Y);
        }

        if (flipV)
        {
            return new SCPoint(this.Width, this.Height);
        }
        
        if (flipH)
        {
            return new SCPoint(this.X, this.Height);
        }
        
        return new SCPoint(this.Width, this.Y);
    }
}