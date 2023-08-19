using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class NoBackground : IImage
{
    private readonly SCSlide slide;
    private readonly Lazy<SCImage> image;

    public NoBackground(SCSlide slide)
    {
        this.slide = slide;
        this.image = new Lazy<SCImage>(this.CreateImage);
    }

    public string MIME => this.image.Value.MIME;

    public Task<byte[]> BinaryData => this.image.Value.BinaryData;

    public string Name => this.image.Value.Name;

    public void SetImage(Stream stream)
    {
        this.image.Value.SetImage(stream);
    }

    public void SetImage(byte[] bytes)
    {
        this.image.Value.SetImage(bytes);
    }

    public void SetImage(string filePath)
    {
        this.image.Value.SetImage(filePath);
    }
    
    private SCImage CreateImage()
    {
        var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
        var pBackground = new P.Background(
            new P.BackgroundProperties(
                new A.BlipFill(
                    new A.Blip { Embed = rId })));
        this.slide.SDKSlidePart.Slide.CommonSlideData!.InsertAt(pBackground, 0);

        var aBlipFill = pBackground.Descendants<A.BlipFill>().SingleOrDefault();
        var picReference = aBlipFill?.Blip?.Embed!;

        var imagePart = this.slide.SDKSlidePart.AddNewPart<ImagePart>("image/png", rId);
        var backgroundImage = new SCImage(imagePart, picReference, this.slide.SDKSlidePart, this.slide.PresentationInternal);

        return backgroundImage;
    }
}