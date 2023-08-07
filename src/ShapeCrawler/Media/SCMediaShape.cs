using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Media;

internal abstract class SCMediaShape : SCShape
{
    private readonly TypedOpenXmlPart slideTypedOpenXmlPart;

    protected SCMediaShape(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf, 
        TypedOpenXmlPart slideTypedOpenXmlPart) 
        : base(pShapeTreeChild, slideOf, shapeCollectionOf)
    {
        this.slideTypedOpenXmlPart = slideTypedOpenXmlPart;
    }

    public byte[] BinaryData => this.GetBinaryData();

    public string MIME => this.GetMime();

    private byte[] GetBinaryData()
    {
        var pPic = (P.Picture)this.PShapeTreeChild;
        var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.slideTypedOpenXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var bytes = stream.ToArray();
        stream.Close();

        return bytes;
    }

    private string GetMime()
    {
        var pPic = (P.Picture)this.PShapeTreeChild;
        var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = this.slideTypedOpenXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

        return relationship.DataPart.ContentType;
    }
}