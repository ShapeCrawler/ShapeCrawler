using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Drawing;

internal sealed class PictureImage : IImage
{
    private readonly SCSlidePicture parentPicture;
    private ImagePart sdkImagePart;
    private readonly DocumentFormat.OpenXml.Drawing.Blip aBlip;

    internal PictureImage(SCSlidePicture slidePicture, DocumentFormat.OpenXml.Drawing.Blip aBlip)
    {
        this.parentPicture = slidePicture;
        this.aBlip = aBlip;
        this.sdkImagePart = (ImagePart)slidePicture.SDKSlidePart().GetPartById(aBlip.Embed!.Value!);
    }
    
    public string MIME => this.sdkImagePart.ContentType;

    public string Name => this.GetName();

    public void Update(Stream stream)
    {
        var imageParts = this.parentPicture.SDKImageParts();
        var isSharedImagePart = imageParts.Count(x=>x == this.sdkImagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.parentPicture.SDKSlidePart().AddNewPart<ImagePart>("image/png", rId);
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