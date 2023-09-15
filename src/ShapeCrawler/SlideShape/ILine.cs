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
    SCPoint StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    SCPoint EndPoint { get; }
}

internal sealed class SlideLine : ILine, IRemoveable
{
    private readonly P.ConnectionShape pConnectionShape;
    private readonly SimpleShape simpleShape;

    internal SlideLine(SlidePart sdkSlidePart, P.ConnectionShape pConnectionShape)
        : this(
            pConnectionShape,
            new SimpleShape(pConnectionShape),
            new SlideShapeOutline(sdkSlidePart, pConnectionShape.ShapeProperties!)
        )
    {
    }

    private SlideLine(
        P.ConnectionShape pConnectionShape,
        SimpleShape simpleShape,
        SlideShapeOutline shapeOutline)
    {
        this.pConnectionShape = pConnectionShape;
        this.simpleShape = simpleShape;
        this.Outline = shapeOutline;
    }

    public SCShapeType ShapeType => SCShapeType.Line;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public SCGeometry GeometryType => SCGeometry.Line;
    public SCPoint StartPoint
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
                return new SCPoint(this.X, this.Y);
            }

            if (flipH)
            {
                return new SCPoint(this.X + this.Width, this.Y);
            }

            return new SCPoint(this.X, this.Y);
        }
    }
    
    public SCPoint EndPoint
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
    
    #region SimpleShape

    public bool IsTextHolder => this.simpleShape.IsTextHolder;
    public ITextFrame TextFrame => this.simpleShape.TextFrame;
    public double Rotation => this.simpleShape.Rotation;
    public ITable AsTable() => this.simpleShape.AsTable();
    public IMediaShape AsMedia() => this.simpleShape.AsMedia();
    public bool HasFill => this.simpleShape.HasFill;
    public IShapeFill Fill => this.simpleShape.Fill;

    public int Width
    {
        get => this.simpleShape.Width;
        set => this.simpleShape.Width = value;
    }

    public int Height
    {
        get => this.simpleShape.Height;
        set => this.simpleShape.Height = value;
    }

    public int Id => this.simpleShape.Id;

    public string Name => this.simpleShape.Name;

    public bool Hidden => this.simpleShape.Hidden;

    public bool IsPlaceholder => this.simpleShape.IsPlaceholder;

    public int X
    {
        get => this.simpleShape.X;
        set => this.simpleShape.X = value;
    }

    public int Y
    {
        get => this.simpleShape.Y;
        set => this.simpleShape.Y = value;
    }

    public IPlaceholder Placeholder => this.simpleShape.Placeholder;

    public string? CustomData
    {
        get => this.simpleShape.CustomData;
        set => this.simpleShape.CustomData = value;
    }

    #endregion SimpleShape
    
    void IRemoveable.Remove()
    {
        this.pConnectionShape.Remove();
    }
}