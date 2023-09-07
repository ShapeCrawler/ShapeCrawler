// ReSharper disable CheckNamespace

using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal record SlideOLEObject : IShape, IRemoveable
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly Shape shape;

    internal SlideOLEObject(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame)
        : this(sdkSlidePart, pGraphicFrame, new Shape(pGraphicFrame))
    {
    }

    private SlideOLEObject(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame, Shape shape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pGraphicFrame = pGraphicFrame;
        this.shape = shape;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, pGraphicFrame.Descendants<P.ShapeProperties>().First(), false);
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

    public IPlaceholder Placeholder => new NullPlaceholder();

    public SCGeometry GeometryType => this.shape.GeometryType();

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.OLEObject;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }
    public bool IsTextHolder => false;

    public ITextFrame TextFrame => throw new SCException(
        $"The OLE object cannot contain text. Use {nameof(IShape.IsTextHolder)} method to check if the shape is a text holder.");
    public double Rotation { get; }
    public ITable AsTable() => throw new SCException($"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

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

    void IRemoveable.Remove()
    {
        this.pGraphicFrame.Remove();
    }
}