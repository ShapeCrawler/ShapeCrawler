using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     Text Shape located on the slide.
/// </summary>
internal sealed class TextRootSlideShape : IRootSlideShape
{
    private readonly IRootSlideShape rootSlideShape;

    internal TextRootSlideShape(SlidePart sdkSlidePart, IRootSlideShape rootSlideShape, P.TextBody pTextBody)
    {
        this.rootSlideShape = rootSlideShape;
        this.TextFrame = new TextFrame(sdkSlidePart, pTextBody);
    }
    
    public bool IsTextHolder => true;

    public ITextFrame TextFrame { get; }
    
    #region RootSlideShape

    public int X
    {
        get => this.rootSlideShape.X;
        set => this.rootSlideShape.X = value;
    }

    public int Y
    {
        get => this.rootSlideShape.Y;
        set => this.rootSlideShape.Y = value;
    }

    public int Width
    {
        get => this.rootSlideShape.Width;
        set => this.rootSlideShape.Width = value;
    }

    public int Height
    {
        get => this.rootSlideShape.Height;
        set => this.rootSlideShape.Height = value;
    }

    public int Id => this.rootSlideShape.Id;
    public string Name => this.rootSlideShape.Name;
    public bool Hidden => this.rootSlideShape.Hidden;

    public SCGeometry GeometryType => this.rootSlideShape.GeometryType;

    public string? CustomData
    {
        get => this.rootSlideShape.CustomData;
        set => this.rootSlideShape.CustomData = value;
    }

    public SCShapeType ShapeType => this.rootSlideShape.ShapeType;
    public bool HasOutline => this.rootSlideShape.HasOutline;
    public IShapeOutline Outline => this.rootSlideShape.Outline;
    public IShapeFill Fill => this.rootSlideShape.Fill;
    public double Rotation => this.rootSlideShape.Rotation;

    public void Duplicate() => this.rootSlideShape.Duplicate();
    
    public ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() => this.rootSlideShape.AsMedia();

    public bool IsPlaceholder => this.rootSlideShape.IsPlaceholder;

    public IPlaceholder Placeholder => this.rootSlideShape.Placeholder;

    #endregion RootSlideShape
}