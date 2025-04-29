using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IMediaShape : IShape
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string Mime { get; }
    
    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();
}

internal class MediaShape : Shape, IMediaShape
{
    private readonly P.Picture pPicture;

    internal MediaShape(P.Picture pPicture)
        : base(pPicture)
    {
        this.pPicture = pPicture;
        this.Outline = new SlideShapeOutline(pPicture.ShapeProperties!);
        this.Fill = new ShapeFill(pPicture.ShapeProperties!);
    }

    public override ShapeContent ShapeContent => ShapeContent.Video;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removable => true;

    public string Mime
    {
        get
        {
            var openXmlPart = this.pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship =
                openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

            return relationship.DataPart.ContentType;
        }
    }

    public byte[] AsByteArray()
    {
        var openXmlPart = this.pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        stream.Close();

        return ms.ToArray();
    }

    public override void Remove() => this.pPicture.Remove();
}