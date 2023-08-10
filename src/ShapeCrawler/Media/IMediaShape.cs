using System;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
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
    public byte[] BinaryData { get; }

    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
}

internal record SCSlideMediaShape : IMediaShape
{
    private readonly P.Picture pPicture;
    private readonly SlidePart sdkSlidePart;
    private readonly Shape shape;
    private readonly Lazy<SCSlidePlaceholder?> placeholder;

    internal SCSlideMediaShape(
        P.Picture pPicture,
        SCSlide slide,
        SCSlideShapes shapes,
        SlidePart sdkSlidePart)
    {
        this.pPicture = pPicture;
        this.sdkSlidePart = sdkSlidePart;
        this.SlideStructure = slide;
        this.shape = new Shape(pPicture);
        this.placeholder = new Lazy<SCSlidePlaceholder?>(this.ParsePlaceholderOrNull);
    }

    private SCSlidePlaceholder? ParsePlaceholderOrNull()
    {
        var pPlaceholder = this.pPicture.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SCSlidePlaceholder(pPlaceholder);
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

    public IPlaceholder? Placeholder => this.placeholder.Value;
    
    public string? CustomData { get; set; }
    
    public SCShapeType ShapeType => SCShapeType.Video;
    
    public ISlideStructure SlideStructure { get; }
    
    public IAutoShape? AsAutoShape()
    {
        return null;
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
    
    public byte[] BinaryData => ParseBinaryData();

    private byte[] ParseBinaryData()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var bytes = stream.ToArray();
        stream.Close();

        return bytes;
    }

    public string MIME => ParseMIME();

    private string ParseMIME()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

        return relationship.DataPart.ContentType;
    }
}