using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed record LayoutShape : IShape
{
    private readonly SlideLayoutPart sdkLayoutPart;
    private readonly P.Shape pShape;
    private readonly Shape shape;
    private readonly Lazy<LayoutShapeFill> autoShapeFill;

    internal LayoutShape(SlideLayoutPart sdkLayoutPart, P.Shape pShape, Shape shape, LayoutShapeOutline outline)
    {
        this.sdkLayoutPart = sdkLayoutPart;
        this.pShape = pShape;
        this.shape = shape;
        this.Outline = outline;
        this.autoShapeFill = new Lazy<LayoutShapeFill>(this.ParseFill);
    }

    private LayoutShapeFill ParseFill()
    {
        var useBgFill = this.pShape.UseBackgroundFill;
        return new LayoutShapeFill(this.sdkLayoutPart, this.pShape.GetFirstChild<P.ShapeProperties>() !, useBgFill!);
    }

    #region Shape Properties

    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden => this.shape.Hidden();

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException(
        $"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill => this.autoShapeFill.Value;

    public bool IsTextHolder => false;
    public ITextFrame TextFrame => new NullTextFrame();
    public double Rotation { get; }
    public ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media.");

    #endregion Shape Properties
}