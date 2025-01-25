using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFillImage : IImage
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Blip aBlip;
    private ImagePart sdkImagePart;

    internal ShapeFillImage(OpenXmlPart sdkTypedOpenXmlPart, A.BlipFill aBlipFill, ImagePart sdkImagePart)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aBlip = aBlipFill.Blip!;
        this.sdkImagePart = sdkImagePart;
    }

    public string Mime => this.sdkImagePart.ContentType;

    public string Name => Path.GetFileName(this.sdkImagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var isSharedImagePart = this.sdkTypedOpenXmlPart.GetPartsOfType<ImagePart>().Count(imagePart => imagePart == this.sdkImagePart) > 1;
        if (isSharedImagePart)
        {            
            var rId = default(RelationshipId).New();
            this.sdkImagePart = this.sdkTypedOpenXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.sdkImagePart.FeedData(stream);
    }

    public byte[] AsByteArray() => new SImagePart(this.sdkImagePart).AsBytes(); 
}