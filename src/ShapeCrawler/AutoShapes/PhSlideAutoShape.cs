using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

/// <summary>
///     The placeholder AutoShape on a slide. 
/// </summary>
internal sealed record PhSlideAutoShape : IAutoShape
{
    private readonly LayoutAutoShape layoutAutoShape;
    private readonly SlideAutoShape slideAutoShape;
    private readonly P.PlaceholderShape pPlaceholderShape;

    internal PhSlideAutoShape(
        SlideAutoShape slideAutoShape, 
        P.PlaceholderShape pPlaceholderShape, 
        LayoutAutoShape layoutAutoShape)
    {
        this.slideAutoShape = slideAutoShape;
        this.pPlaceholderShape = pPlaceholderShape;
        this.layoutAutoShape = layoutAutoShape;
    }
    
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.slideAutoShape.Id;
    public string Name => this.slideAutoShape.Name;
    public bool Hidden => this.slideAutoShape.Hidden;
    public bool IsPlaceholder() => true;

    public IPlaceholder Placeholder => new SlidePlaceholder(this.pPlaceholderShape);
    public SCGeometry GeometryType => this.slideAutoShape.GeometryType;
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => this.slideAutoShape.ShapeType;
    public IAutoShape AsAutoShape() => this;

    public IShapeOutline Outline => this.slideAutoShape.Outline;
    public IShapeFill Fill => this.slideAutoShape.Fill;
    public ITextFrame TextFrame => this.slideAutoShape.TextFrame;
    public bool IsTextHolder() => this.slideAutoShape.IsTextHolder();

    public double Rotation => this.slideAutoShape.Rotation;
}