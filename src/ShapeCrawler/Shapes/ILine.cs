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

internal sealed record SCLine : ILine
{
    private readonly P.ConnectionShape pConnectionShape;
    private readonly SlideShapes shapeCollection;
    private readonly Shape shape;

    internal SCLine(
        P.ConnectionShape pConnectionShape,
        SlideShapes parentShapeCollection,
        Shape shape)
    {
        this.pConnectionShape = pConnectionShape;
        this.shapeCollection = parentShapeCollection;
        this.shape = shape;
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

    public bool IsPlaceholder()
    {
        return false; 
    }

    public IPlaceholder Placeholder => new NullPlaceholder($"A line cannot be a placeholder. Use {nameof(IShape.Placeholder)} to check if the shape is a placeholder.");
    public SCGeometry GeometryType => SCGeometry.Line;

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }
    public SCShapeType ShapeType => SCShapeType.Line;
    public IAutoShape AsAutoShape()
    {
        return this;
    }

    public ITextFrame? TextFrame => null;
    public bool IsTextHolder()
    {
        throw new System.NotImplementedException();
    }

    public double Rotation { get; }

    public IAutoShape Duplicate()
    {
        throw new System.NotImplementedException();
    }

    public IShapeOutline Outline { get; }
    public IShapeFill Fill => new SCNullShapeFill();
    
    public SCPoint StartPoint => this.GetStartPoint();
    
    public SCPoint EndPoint => this.GetEndPoint();

    internal void Draw(SKCanvas canvas)
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

    public int X { get; set; }
    public int Y { get; set; }
}