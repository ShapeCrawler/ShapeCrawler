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

internal class SlideMediaShape : Shape, IMediaShape, IRemoveable
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.Picture pPicture;

    internal SlideMediaShape(SlidePart sdkSlidePart, P.Picture pPicture)
    :base(pPicture)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pPicture = pPicture;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pPicture.ShapeProperties!);
        this.Fill = new SlideShapeFill(sdkSlidePart, pPicture.ShapeProperties!, false);
    }
    public override SCShapeType ShapeType => SCShapeType.Video;
    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override bool HasFill => true;
    public override IShapeFill Fill { get; }
    
    public string MIME
    {
        get
        {
            var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship = this.sdkSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
            
            return relationship.DataPart.ContentType;
        }
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
    
    void IRemoveable.Remove() => this.pPicture.Remove();
}