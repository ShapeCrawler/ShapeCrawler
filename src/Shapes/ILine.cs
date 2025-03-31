using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine : IShape
{
    /// <summary>
    ///    Gets the start point of the line.
    /// </summary>
    Point StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    Point EndPoint { get; }
}

internal sealed class SlideLine : Shape, ILine
{
    private readonly P.ConnectionShape pConnectionShape;

    internal SlideLine(P.ConnectionShape pConnectionShape)
        : this(pConnectionShape, new SlideShapeOutline(pConnectionShape.ShapeProperties!))
    {
    }

    private SlideLine(P.ConnectionShape pConnectionShape, SlideShapeOutline shapeOutline)
        : base(pConnectionShape)
    {
        this.pConnectionShape = pConnectionShape;
        this.Outline = shapeOutline;
    }

    public override ShapeContent ShapeContent => ShapeContent.Line;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override Geometry GeometryType => Geometry.Line;

    public Point StartPoint
    {
        get
        {
            var aTransform2D = this.pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (flipH && (this.Height == 0 || flipV))
            {
                return new Point(this.X, (decimal)this.Y);
            }

            if (flipH)
            {
                return new Point((decimal)(this.X + this.Width), (decimal)this.Y);
            }

            return new Point(this.X, (decimal)this.Y);
        }
    }

    public Point EndPoint
    {
        get
        {
            var aTransform2D = this.pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (this.Width == 0)
            {
                return new Point((decimal)this.X, (decimal)this.Height);
            }

            if (flipH && this.Height == 0)
            {
                return new Point((decimal)(this.X - this.Width), (decimal)this.Y);
            }

            if (flipV)
            {
                return new Point((decimal)this.Width, (decimal)this.Height);
            }

            if (flipH)
            {
                return new Point((decimal)this.X, (decimal)this.Height);
            }

            return new Point((decimal)this.Width, (decimal)this.Y);
        }
    }

    public override bool Removeable => true;
    
    public override void Remove() => this.pConnectionShape.Remove();
}