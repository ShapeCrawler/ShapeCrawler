using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class SlidePictureImage : IImage
{
    private readonly OpenXmlPart openXmlPart;
    private readonly A.Blip aBlip;
    private ImagePart imagePart;

    internal SlidePictureImage(A.Blip aBlip)
    {
        this.aBlip = aBlip;
        this.openXmlPart = aBlip.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        this.imagePart = (ImagePart)this.openXmlPart.GetPartById(aBlip.Embed!.Value!);
    }

    public string Mime => this.imagePart.ContentType;

    public string Name => Path.GetFileName(this.imagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var presDocument = (PresentationDocument)this.openXmlPart.OpenXmlPackage;
        var slideParts = presDocument.PresentationPart!.SlideParts;
        var allABlips = slideParts.SelectMany(slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!.Descendants<A.Blip>());
        var isSharedImagePart = allABlips.Count(blip => blip.Embed!.Value == this.aBlip.Embed!.Value) > 1;
        if (isSharedImagePart)
        {
            var rId = RelationshipId.New();
            this.imagePart = this.openXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.imagePart.FeedData(stream);
    }

    public byte[] AsByteArray() => new SCImagePart(this.imagePart).AsBytes();
}