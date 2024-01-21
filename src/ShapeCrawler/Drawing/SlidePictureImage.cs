using System;
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
    
    public string MIME => this.sdkImagePart.ContentType;

    public string Name => Path.GetFileName(this.sdkImagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var sdkPresDocument = (PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage;
        var presSdkSlideParts = sdkPresDocument.PresentationPart!.SlideParts;
        var allABlip = presSdkSlideParts.SelectMany(x => x.Slide.CommonSlideData!.ShapeTree!.Descendants<A.Blip>());
        var isSharedImagePart = allABlip.Count(x => x.Embed!.Value == this.aBlip.Embed!.Value) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.sdkTypedOpenXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.sdkImagePart.FeedData(stream);
    }

    public void Update(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);

        this.Update(stream);
    }

    public void Update(string file)
    {
        byte[] sourceBytes = File.ReadAllBytes(file);
        this.Update(sourceBytes);
    }
    
    public byte[] AsByteArray()
    {
        var stream = this.sdkImagePart.GetStream();
        var bytes = new byte[stream.Length];
        stream.Read(bytes, 0, (int)stream.Length);
        stream.Close();
        
        return bytes;
    }
}