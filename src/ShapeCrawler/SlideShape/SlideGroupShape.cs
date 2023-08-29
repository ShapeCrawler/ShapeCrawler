using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Shape = ShapeCrawler.Shapes.Shape;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.SlideShape;

internal sealed record SlideGroupShape : IGroupShape, IRemoveable
{
    private readonly P.GroupShape sdkPGroupShape;
    private readonly Shape shape;

    internal SlideGroupShape(SlidePart sdkSlidePart, P.GroupShape sdkPGroupShape)
    {
        this.sdkPGroupShape = sdkPGroupShape;
        this.shape = new Shape(sdkPGroupShape);
        this.Shapes = new SlideGroupedShapes(sdkSlidePart, sdkPGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(sdkSlidePart, sdkPGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, sdkPGroupShape.Descendants<P.ShapeProperties>().First(), false);
    }

    public IReadOnlyShapeCollection Shapes { get; }

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

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException($"Group Shape cannot be a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");
    public SCGeometry GeometryType { get; }

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.Group;
    public bool HasOutline { get; }
    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }
    public bool IsTextHolder => false;
    public ITextFrame TextFrame => throw new SCException($"Group Shape cannot be a text holder. Use {nameof(IShape.IsTextHolder)} property to check if the shape is a text holder.");
    public double Rotation { get; }

    internal void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new System.NotImplementedException();
    }

    internal string ToJson()
    {
        throw new System.NotImplementedException();
    }

    public int X
    {
        get => this.shape.X();
        set => this.shape.UpdateX(value);
    }

    public int Y
    {
        get => this.shape.Y();
        set => this.shape.UpdateY(value);
    }

    void IRemoveable.Remove()
    {
        this.sdkPGroupShape.Remove();
    }
}