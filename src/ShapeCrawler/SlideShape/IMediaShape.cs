using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IMediaShape : IShape
{
    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();

    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
}

internal record SlideMediaShape : IMediaShape, IRemoveable
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.Picture pPicture;
    private readonly Shape shape;

    internal SlideMediaShape(SlidePart sdkSlidePart, P.Picture pPicture)
        : this(sdkSlidePart, pPicture, new Shape(pPicture))
    {
    }

    private SlideMediaShape(SlidePart sdkSlidePart, P.Picture pPicture, Shape shape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pPicture = pPicture;
        this.shape = shape;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pPicture.ShapeProperties!);
        this.Fill = new SlideShapeFill(sdkSlidePart, pPicture.ShapeProperties!, false);
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

    public SCGeometry GeometryType => this.shape.GeometryType();

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => new NullPlaceholder();

    public string? CustomData { get; set; }

    public SCShapeType ShapeType => SCShapeType.Video;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }

    public bool IsTextHolder => false;

    public ITextFrame TextFrame =>
        throw new SCException(
            $"The media shape is not a text holder. Use {nameof(IShape.IsTextHolder)} property to check if the shape is a text holder.");
    public double Rotation { get; }
    public ITable AsTable() => throw new SCException($"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() => this;

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

    public byte[] AsByteArray()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var bytes = stream.ToArray();
        stream.Close();

        return bytes;
    }

    public string MIME => ParseMIME();

    private string ParseMIME()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

        return relationship.DataPart.ContentType;
    }

    void IRemoveable.Remove()
    {
        this.pPicture.Remove();
    }
}