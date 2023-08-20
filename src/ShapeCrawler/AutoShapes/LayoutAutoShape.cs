using ShapeCrawler.Shapes;

namespace ShapeCrawler.AutoShapes;

internal sealed record LayoutAutoShape : IAutoShape
{
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id { get; }
    public string Name { get; }
    public bool Hidden { get; }
    public bool IsPlaceholder()
    {
        throw new System.NotImplementedException();
    }

    public IPlaceholder Placeholder => new NullPlaceholder();
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public IAutoShape AsAutoShape()
    {
        throw new System.NotImplementedException();
    }

    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }
    public ITextFrame TextFrame { get; }
    public bool IsTextHolder()
    {
        throw new System.NotImplementedException();
    }

    public double Rotation { get; }
}