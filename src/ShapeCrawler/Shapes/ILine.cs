using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
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
    SCPoint StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    SCPoint EndPoint { get; }
}

internal sealed record SlideLine : ILine, IRemoveable
{
    private readonly P.ConnectionShape pConnectionShape;
    private readonly Shape shape;

    internal SlideLine(SlidePart sdkSlidePart, P.ConnectionShape pConnectionShape)
        : this(
            pConnectionShape,
            new Shape(pConnectionShape),
            new SlideShapeOutline(sdkSlidePart, pConnectionShape.ShapeProperties!)
        )
    {
    }

    private SlideLine(
        P.ConnectionShape pConnectionShape,
        Shape shape,
        SlideShapeOutline shapeOutline)
    {
        this.pConnectionShape = pConnectionShape;
        this.shape = shape;
        this.Outline = shapeOutline;
    }

    public int Width
    {
        get => this.shape.Width();
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height();
        set => this.shape.UpdateHeight(value);
    }

    public int Id => this.shape.Id();

    public string Name => this.shape.Name();

    public bool Hidden => this.shape.Hidden();

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder =>
        new NullPlaceholder(
            $"A line cannot be a placeholder. Use {nameof(IShape.Placeholder)} to check if the shape is a placeholder.");

    public SCGeometry GeometryType => SCGeometry.Line;

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.Line;

    public bool IsTextHolder => false;
    public ITextFrame TextFrame => new NullTextFrame();

    public double Rotation { get; }

    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill => new SCNullShapeFill();

    public SCPoint StartPoint => this.ParseStartPoint();

    public SCPoint EndPoint => this.ParseEndPoint();

    internal void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    private SCPoint ParseStartPoint()
    {
        var aTransform2D = this.pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
        var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
        var flipH = horizontalFlip != null && horizontalFlip.Value;
        var verticalFlip = aTransform2D.VerticalFlip?.Value;
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

    private SCPoint ParseEndPoint()
    {
        var aTransform2D = this.pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
        var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
        var flipH = horizontalFlip != null && horizontalFlip.Value;
        var verticalFlip = aTransform2D.VerticalFlip?.Value;
        var flipV = verticalFlip != null && verticalFlip.Value;

        if (this.Width == 0)
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

    public int X { get; set; }
    public int Y { get; set; }
    
    void IRemoveable.Remove()
    {
        this.pConnectionShape.Remove();
    }
}