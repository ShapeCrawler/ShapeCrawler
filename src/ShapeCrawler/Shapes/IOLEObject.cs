// ReSharper disable CheckNamespace

using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using SkiaSharp;

namespace ShapeCrawler;

internal record SCSlideOLEObject : IShape
{
    private readonly Shape shape;

    internal SCSlideOLEObject(
        OpenXmlCompositeElement pShapeTreeChild, 
        SCSlide slide,
        SCSlideShapes shapes)
    {
        this.shape = new Shape(pShapeTreeChild);
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
    public IPlaceholder? Placeholder => null;

    public SCGeometry GeometryType => this.shape.GeometryType();
    public string? CustomData 
    {
        get => this.shape.CustomData(); 
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.OLEObject;
    public ISlideStructure SlideStructure { get; }
    public IAutoShape? AsAutoShape()
    {
        throw new System.NotImplementedException();
    }

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
}