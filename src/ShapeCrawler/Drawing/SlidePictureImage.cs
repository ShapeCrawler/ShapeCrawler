using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class SlidePictureImage : IImage
{
    private ImagePart sdkImagePart;
    private readonly SlidePart sdkSlidePart;
    private readonly A.Blip aBlip;

    internal SlidePictureImage(SlidePart sdkSlidePart, A.Blip aBlip)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aBlip = aBlip;
        this.sdkImagePart = (ImagePart)sdkSlidePart.GetPartById(aBlip.Embed!.Value!);
    }
    
    public string MIME => this.sdkImagePart.ContentType;

    public string Name => this.GetName();

    public void Update(Stream stream)
    {
        var sdkPresDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
        var presSdkSlideParts = sdkPresDocument.PresentationPart!.SlideParts;
        var allABlip = presSdkSlideParts.SelectMany(x => x.Slide.CommonSlideData!.ShapeTree!.Descendants<A.Blip>());
        var isSharedImagePart = allABlip.Count(x => x.Embed!.Value == this.aBlip.Embed!.Value) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.sdkSlidePart.AddNewPart<ImagePart>("image/png", rId);
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

    public void Update(string filePath)
    {
        byte[] sourceBytes = File.ReadAllBytes(filePath);
        this.Update(sourceBytes);
    }
    
    private string GetName()
    {
        return Path.GetFileName(this.sdkImagePart.Uri.ToString());
    }

    public byte[] BinaryData()
    {
        var stream = this.sdkImagePart.GetStream();
        var bytes = new byte[stream.Length];
        stream.Read(bytes, 0, (int)stream.Length);
        stream.Close();
        
        return bytes;
    }
}