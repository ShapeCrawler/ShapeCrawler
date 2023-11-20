using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape containing video content.
/// </summary>
public interface IMediaShape : IShape
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }
    
    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();

}

internal class MediaShape : Shape, IMediaShape
{
    private readonly P.Picture pPicture;

    internal MediaShape(TypedOpenXmlPart sdkTypedOpenXmlPart, P.Picture pPicture)
        : base(sdkTypedOpenXmlPart, pPicture)
    {
        this.pPicture = pPicture;
        this.Outline = new SlideShapeOutline(sdkTypedOpenXmlPart, pPicture.ShapeProperties!);
        this.Fill = new ShapeFill(sdkTypedOpenXmlPart, pPicture.ShapeProperties!);
    }

    public override ShapeType ShapeType => ShapeType.Video;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;

    public string MIME
    {
        get
        {
            var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship =
                this.sdkTypedOpenXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

            return relationship.DataPart.ContentType;
        }
    }

    public byte[] AsByteArray()
    {
        var p14Media = this.pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.sdkTypedOpenXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        stream.Close();

        return ms.ToArray();
    }

    public override void Remove() => this.pPicture.Remove();
}