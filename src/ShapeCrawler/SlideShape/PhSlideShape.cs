using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     The placeholder AutoShape on a slide. 
/// </summary>
internal sealed record PhSlideShape : IShape
{
    private readonly LayoutShape layoutShape;
    private readonly SlideShape slideShape;
    private readonly P.PlaceholderShape sdkPPlaceholderShape;

    internal PhSlideShape(
        SlideShape slideShape, 
        P.PlaceholderShape sdkPPlaceholderShape, 
        LayoutShape layoutShape)
    {
        this.slideShape = slideShape;
        this.sdkPPlaceholderShape = sdkPPlaceholderShape;
        this.layoutShape = layoutShape;
    }
    
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.slideShape.Id;
    public string Name => this.slideShape.Name;
    public bool Hidden => this.slideShape.Hidden;
    public bool IsPlaceholder => true;
    public IPlaceholder Placeholder => new SlidePlaceholder(this.sdkPPlaceholderShape);
    public SCGeometry GeometryType => this.slideShape.GeometryType;
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => this.slideShape.ShapeType;
    public bool HasOutline => this.slideShape.HasOutline;
    public IShapeOutline Outline => this.slideShape.Outline;
    public IShapeFill Fill => this.slideShape.Fill;
    public bool IsTextHolder { get; }
    public ITextFrame TextFrame => this.slideShape.TextFrame;

    public double Rotation => this.slideShape.Rotation;
}