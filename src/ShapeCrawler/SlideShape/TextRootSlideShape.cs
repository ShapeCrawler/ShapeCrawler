using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     Text Shape located on the slide.
/// </summary>
internal sealed record TextRootSlideShape : IRootSlideShape
{
    private readonly IRootSlideShape rootSlideShape;

    internal TextRootSlideShape(SlidePart sdkSlidePart, IRootSlideShape rootSlideShape, P.TextBody pTextBody)
    {
        this.rootSlideShape = rootSlideShape;
        this.TextFrame = new TextFrame(sdkSlidePart, pTextBody);
    }

    #region Root Slide Shape

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

    #endregion Root Slide Shape
    
    public bool IsTextHolder => true;

    public ITextFrame TextFrame { get; }
    
    public ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media.");

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException(
        $"Text shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if the shape is a placeholder.");
    
}