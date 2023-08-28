using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record LayoutAutoShape : IAutoShape
{
    private readonly SlideLayoutPart sdkLayoutPart;
    private readonly P.Shape pShape;
    private readonly Shape shape;
    private readonly Lazy<LayoutAutoShapeFill> autoShapeFill;

    internal LayoutAutoShape(SlideLayoutPart sdkLayoutPart, P.Shape pShape, Shape shape, LayoutShapeOutline outline)
    {
        this.sdkLayoutPart = sdkLayoutPart;
        this.pShape = pShape;
        this.shape = shape;
        this.Outline = outline;
        this.autoShapeFill = new Lazy<LayoutAutoShapeFill>(this.ParseFill);
        this.Placeholder = new NullPlaceholder();
    }

    private LayoutAutoShapeFill ParseFill()
    {
        var useBgFill = pShape.UseBackgroundFill;
        return new LayoutAutoShapeFill(this.sdkLayoutPart, this.pShape.GetFirstChild<P.ShapeProperties>() !, useBgFill!);
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