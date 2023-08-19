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

    public Task<byte[]> BinaryData => this.GetBinaryData();

    public string Name => this.GetName();

    public void UpdateImage(Stream stream)
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

    private async Task<byte[]> GetBinaryData()
    {
        var stream = this.sdkImagePart.GetStream();
        var bytes = new byte[stream.Length];
        await stream.ReadAsync(bytes, 0, (int)stream.Length).ConfigureAwait(false);
        stream.Close();

        return bytes;
    }
}