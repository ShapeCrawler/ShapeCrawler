using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record LayoutAutoShape : IAutoShape
{
    private readonly P.Shape pShape;
    private readonly LayoutShapes parentShapeCollection;
    private readonly Shape shape;

    internal LayoutAutoShape(P.Shape pShape, LayoutShapes parentShapeCollection, Shape shape)
    {
        this.pShape = pShape;
        this.parentShapeCollection = parentShapeCollection;
        this.shape = shape;
    }

    #region Shape Properties
    
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden => this.shape.Hidden();
    public bool IsPlaceholder() => false;
    public IPlaceholder Placeholder => new NullPlaceholder();
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public IAutoShape AsAutoShape() => this;
    public IShapeOutline Outline => new ShapeOutline(this.parentShapeCollection.SlideMaster(), this.pShape.ShapeProperties!);
    public IShapeFill Fill { get; }
    public ITextFrame TextFrame => new NullTextFrame();
    public bool IsTextHolder() => false;
    public double Rotation { get; }
    
    #endregion Shape Properties
}