using System;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record LayoutAutoShape : IAutoShape
{
    private readonly P.Shape pShape;
    private readonly Shape shape;
    private readonly Lazy<SlideAutoShapeFill> autoShapeFill;

    internal LayoutAutoShape(P.Shape pShape, Shape shape, LayoutShapeOutline outline)
    {
        this.pShape = pShape;
        this.shape = shape;
        this.autoShapeFill = new Lazy<SlideAutoShapeFill>(this.ParseFill);
        this.Placeholder = new NullPlaceholder();
        this.Outline = outline;
    }

    private SlideAutoShapeFill ParseFill()
    {
        var useBgFill = pShape.UseBackgroundFill;
        return new SlideAutoShapeFill(this.pShape.GetFirstChild<P.ShapeProperties>() !, useBgFill);
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
    public IPlaceholder Placeholder { get; }
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public IAutoShape AsAutoShape() => this;

    public IShapeOutline Outline { get; }

    public IShapeFill Fill => this.autoShapeFill.Value;
    public ITextFrame TextFrame => new NullTextFrame();
    public bool IsTextHolder() => false;
    public double Rotation { get; }

    #endregion Shape Properties
}