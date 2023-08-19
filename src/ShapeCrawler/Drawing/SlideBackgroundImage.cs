using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SlideBackgroundImage : IImage
{
    private ImagePart sdkImagePart;
    private readonly A.Blip aBlip;
    private readonly SCSlide slide;

    public string MIME => this.sdkImagePart.ContentType;

    public Task<byte[]> BinaryData => this.GetBinaryData();

    public string Name => this.GetName();

    public void UpdateImage(Stream stream)
    {
        var imageParts = this.slide.SDKImageParts();
        var isSharedImagePart = imageParts.Count(x=>x == this.sdkImagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.slide.SDKSlidePart().AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.sdkImagePart.FeedData(stream);
    }

    public void SetImage(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);

        this.UpdateImage(stream);
    }

    public void SetImage(string filePath)
    {
        byte[] sourceBytes = File.ReadAllBytes(filePath);
        this.SetImage(sourceBytes);
    }

    internal SlideBackgroundImage (SCSlide slide, A.Blip aBlip, ImagePart sdkImagePart)
    {
        this.slide = slide;
        this.sdkImagePart = sdkImagePart;
        this.aBlip = aBlip;
    }

    private string GetName()
    {
        return Path.GetFileName(this.sdkImagePart.Uri.ToString());
    }

    private async Task<byte[]> GetBinaryData()
    {
        var stream = this.sdkImagePart.GetStream();
        var bytes = new byte[stream.Length];
        await stream.ReadAsync(bytes, 0, (int)stream.Length).ConfigureAwait(false);
        stream.Close();
        
        return bytes;
    }
}