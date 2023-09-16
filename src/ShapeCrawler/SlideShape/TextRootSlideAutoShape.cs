using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     Text Shape located on the slide.
/// </summary>
internal sealed class TextRootSlideAutoShape : IRootSlideAutoShape
{
    private readonly IRootSlideAutoShape rootSlideAutoShape;

    internal TextRootSlideAutoShape(SlidePart sdkSlidePart, IRootSlideAutoShape rootSlideAutoShape, P.TextBody pTextBody)
    {
        this.rootSlideAutoShape = rootSlideAutoShape;
        this.TextFrame = new TextFrame(sdkSlidePart, pTextBody);
    }
    
    public bool IsTextHolder => true;

    public ITextFrame TextFrame { get; }
    
    #region RootSlideShape
    
    public int X
    {
        get => this.rootSlideAutoShape.X;
        set => this.rootSlideAutoShape.X = value;
    }

    public int Y
    {
        get => this.rootSlideAutoShape.Y;
        set => this.rootSlideAutoShape.Y = value;
    }

    public int Width
    {
        get => this.rootSlideAutoShape.Width;
        set => this.rootSlideAutoShape.Width = value;
    }

    public int Height
    {
        get => this.rootSlideAutoShape.Height;
        set => this.rootSlideAutoShape.Height = value;
    }

    public int Id => this.rootSlideAutoShape.Id;
    public string Name => this.rootSlideAutoShape.Name;
    public bool Hidden => this.rootSlideAutoShape.Hidden;

    public SCGeometry GeometryType => this.rootSlideAutoShape.GeometryType;

    public string? CustomData
    {
        get => this.rootSlideAutoShape.CustomData;
        set => this.rootSlideAutoShape.CustomData = value;
    }

    public SCShapeType ShapeType => this.rootSlideAutoShape.ShapeType;
    public bool HasOutline => this.rootSlideAutoShape.HasOutline;
    public IShapeOutline Outline => this.rootSlideAutoShape.Outline;
    public bool HasFill { get; }
    public IShapeFill Fill => this.rootSlideAutoShape.Fill;
    public double Rotation => this.rootSlideAutoShape.Rotation;

    public void Duplicate() => this.rootSlideAutoShape.Duplicate();
    
    public ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() => this.rootSlideAutoShape.AsMedia();

    public bool IsPlaceholder => this.rootSlideAutoShape.IsPlaceholder;

    public IPlaceholder Placeholder => this.rootSlideAutoShape.Placeholder;

    #endregion RootSlideShape
}