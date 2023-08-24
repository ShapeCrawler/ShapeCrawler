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

internal record SlideMediaShape : IMediaShape
{
    private readonly P.Picture pPicture;
    private readonly SlideShapes parentShapeCollection;
    private readonly Shape shape;

    internal SlideMediaShape(P.Picture pPicture, SlideShapes parentShapeCollection, Shape shape)
    {
        this.pPicture = pPicture;
        this.parentShapeCollection = parentShapeCollection;
        this.shape = shape;
    }

    private SlidePlaceholder? ParsePlaceholderOrNull()
    {
        var pPlaceholder = this.pPicture.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SlidePlaceholder(pPlaceholder);
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

    public bool IsPlaceholder() => false;

    public IPlaceholder Placeholder => new NullPlaceholder();
    
    public string? CustomData { get; set; }
    
    public SCShapeType ShapeType => SCShapeType.Video;
    
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
        var relationship = this.parentShapeCollection.SDKSlidePart().DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var bytes = stream.ToArray();
        stream.Close();

        return bytes;
    }

    public string MIME => ParseMIME();

    private string ParseMIME()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.parentShapeCollection.SDKSlidePart().DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

        return relationship.DataPart.ContentType;
    }
}