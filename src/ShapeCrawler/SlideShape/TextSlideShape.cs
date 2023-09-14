using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class TextSlideShape : IShape
{
    private readonly IShape shape;

    internal TextSlideShape(SlidePart sdkSlidePart, IShape shape, P.TextBody pTextBody)
    {
        this.shape = shape;
        this.TextFrame = new TextFrame(sdkSlidePart, pTextBody);
    }

    public bool IsTextHolder => true;
    public ITextFrame TextFrame { get; }
    
    #region IShape

    public int X
    {
        get => this.shape.X;
        set => this.shape.X = value;
    }

    public int Y
    {
        get => this.shape.Y;
        set => this.shape.Y = value;
    }

    public int Width
    {
        get => this.shape.Width;
        set => this.shape.Width = value;
    }

    public int Height
    {
        get => this.shape.Height;
        set => this.shape.Height = value;
    }

    public int Id => this.shape.Id;
    public string Name => this.shape.Name;
    public bool Hidden => this.shape.Hidden;
    public bool IsPlaceholder => this.shape.IsPlaceholder;
    public IPlaceholder Placeholder => this.shape.Placeholder;
    public SCGeometry GeometryType => this.shape.GeometryType;

    public string? CustomData
    {
        get => this.shape.CustomData; 
        set => this.shape.CustomData = value;
    }
    public SCShapeType ShapeType => this.shape.ShapeType;
    public bool HasOutline => this.shape.HasOutline;
    public IShapeOutline Outline => this.shape.Outline;
    public IShapeFill Fill => this.shape.Fill;
    public double Rotation => this.shape.Rotation;
    public ITable AsTable() => this.shape.AsTable();
    public IMediaShape AsMedia() => this.shape.AsMedia();

    #endregion IShape
}