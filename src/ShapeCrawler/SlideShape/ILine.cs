using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

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

internal sealed class SlideLine : Shape, ILine, IRemoveable
{
    private readonly P.ConnectionShape pConnectionShape;

    internal SlideLine(SlidePart sdkSlidePart, P.ConnectionShape pConnectionShape)
        : this(
            pConnectionShape,
            new SlideShapeOutline(sdkSlidePart, pConnectionShape.ShapeProperties!)
        )
    {
    }

    private SlideLine(
        P.ConnectionShape pConnectionShape,
        SlideShapeOutline shapeOutline)
        : base(pConnectionShape)
    {
        this.pConnectionShape = pConnectionShape;
        this.Outline = shapeOutline;
    }

    public override SCShapeType ShapeType => SCShapeType.Line;
    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override SCGeometry GeometryType => SCGeometry.Line;

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
                return new Point(this.X, this.Y);
            }

            if (flipH)
            {
                return new Point(this.X + this.Width, this.Y);
            }

            return new Point(this.X, this.Y);
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
                return new Point(this.X, this.Height);
            }

            if (flipH && this.Height == 0)
            {
                return new Point(this.X - this.Width, this.Y);
            }

            if (flipV)
            {
                return new Point(this.Width, this.Height);
            }

            if (flipH)
            {
                return new Point(this.X, this.Height);
            }

            return new Point(this.Width, this.Y);
        }
    }

    void IRemoveable.Remove()
    {
        this.pConnectionShape.Remove();
    }
}