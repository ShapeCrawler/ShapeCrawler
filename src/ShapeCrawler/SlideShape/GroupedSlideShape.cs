using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.SlideShape;

internal sealed record GroupedSlideShape : IShape
{
    private readonly SlideShape slideShape;

    internal GroupedSlideShape(SlideShape slideShape)
    {
        this.slideShape = slideShape;
    }

    #region Slide APIs
    
    public int X
    {
        get => this.slideShape.X;
        set => this.slideShape.X = value;
    }

    public int Y 
    {
        get => this.slideShape.Y; 
        set => this.slideShape.Y = value;
    }
    
    public int Width
    {
        get => this.slideShape.Width; 
        set => this.slideShape.Width = value;
    }
    public int Height
    {
        get => this.slideShape.Height; 
        set => this.slideShape.Height = value;
    }
    public int Id => this.slideShape.Id;
    public string Name => this.slideShape.Name;
    public bool Hidden => this.slideShape.Hidden;
    
    public string? CustomData
    {
        get => this.slideShape.CustomData; 
        set => this.slideShape.CustomData = value;
    }
    public SCShapeType ShapeType => this.slideShape.ShapeType;
    public bool HasOutline => this.slideShape.HasOutline;
    public IShapeOutline Outline => this.slideShape.Outline;
    public IShapeFill Fill => this.slideShape.Fill;
    
    public double Rotation => this.slideShape.Rotation;
    
    public SCGeometry GeometryType => this.slideShape.GeometryType;
    
    #endregion Slide APIs

    public bool IsPlaceholder => false;
    public IPlaceholder Placeholder =>
        throw new SCException($"Grouped Shape cannot be a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");

    public bool IsTextHolder => false;
    public ITextFrame TextFrame =>
        throw new SCException($"The shape is not a text holder. Use {nameof(IShape.IsTextHolder)} property to check if the shape is a text holder.");

    public ITable AsTable() => throw new SCException($"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");
    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media.");
}