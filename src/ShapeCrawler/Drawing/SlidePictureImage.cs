using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class SlidePictureImage : IImage
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Blip aBlip;
    private ImagePart sdkImagePart;

    internal SlidePictureImage(OpenXmlPart sdkTypedOpenXmlPart, A.Blip aBlip)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aBlip = aBlip;
        this.sdkImagePart = (ImagePart)this.sdkTypedOpenXmlPart.GetPartById(aBlip.Embed!.Value!);
    }

    public string Mime => this.sdkImagePart.ContentType;

    public string Name => Path.GetFileName(this.sdkImagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var sdkPresDocument = (PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage;
        var slideParts = sdkPresDocument.PresentationPart!.SlideParts;
        var allABlips = slideParts.SelectMany(slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!.Descendants<A.Blip>());
        
        var isSharedImagePart = allABlips.Count(blip => blip.Embed!.Value == this.aBlip.Embed!.Value) > 1;
        if (isSharedImagePart)
        {
            var rId = default(RelationshipId).New();
            this.sdkImagePart = this.sdkTypedOpenXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.sdkImagePart.FeedData(stream);
    }

    public byte[] AsByteArray() => new WrappedImagePart(this.sdkImagePart).AsBytes();
}