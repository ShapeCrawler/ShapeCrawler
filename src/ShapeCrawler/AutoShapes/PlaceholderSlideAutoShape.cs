using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record PlaceholderSlideAutoShape : ISlideAutoShape
{
    private readonly LayoutAutoShape layoutAutoShape;
    private readonly SlideAutoShape slideAutoShape;
    private readonly P.PlaceholderShape pPlaceholderShape;

    internal PlaceholderSlideAutoShape(
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
    public int Id { get; }
    public string Name => this.slideAutoShape.Name;
    public bool Hidden { get; }
    public bool IsPlaceholder()
    {
        throw new System.NotImplementedException();
    }

    public IPlaceholder Placeholder => new SlidePlaceholder(this.pPlaceholderShape);
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public IAutoShape AsAutoShape()
    {
        throw new System.NotImplementedException();
    }

    public IShapeOutline Outline => this.slideAutoShape.Outline;
    public IShapeFill Fill => this.slideAutoShape.Fill;
    public ITextFrame TextFrame => this.slideAutoShape.TextFrame;
    public bool IsTextHolder()
    {
        throw new System.NotImplementedException();
    }

    public double Rotation { get; }
    public void Duplicate()
    {
        throw new System.NotImplementedException();
    }
}