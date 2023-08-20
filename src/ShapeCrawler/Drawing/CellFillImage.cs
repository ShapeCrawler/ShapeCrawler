using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class CellFillImage : IImage
{
    private ImagePart sdkImagePart;
    private readonly TableCellFill parentTableCellFill;
    private readonly A.Blip aBlip;

    public string MIME => this.sdkImagePart.ContentType;

    public string Name => this.GetName();

    public void Update(Stream stream)
    {
        List<ImagePart> imageParts = this.parentTableCellFill.SDKImageParts();
        var isSharedImagePart = imageParts.Count(x => x == this.sdkImagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.parentTableCellFill.SDKSlidePart().AddNewPart<ImagePart>("image/png", rId);
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

    internal CellFillImage(A.BlipFill aBlipFill, ImagePart sdkImagePart, TableCellFill parentTableCellFill)
    {
        this.aBlip = aBlipFill.Blip!;
        this.sdkImagePart = sdkImagePart;
        this.parentTableCellFill = parentTableCellFill;
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