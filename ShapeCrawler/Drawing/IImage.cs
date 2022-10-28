using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents an image.
/// </summary>
public interface IImage
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }

    /// <summary>
    ///     Gets binary content.
    /// </summary>
    Task<byte[]> BinaryData { get; }

    /// <summary>
    ///     Gets file name of internal resource.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Sets image with stream.
    /// </summary>
    void SetImage(Stream stream);

    /// <summary>
    ///     Sets image with byte array.
    /// </summary>
    void SetImage(byte[] bytes);

    /// <summary>
    ///     Sets image by specified file path.
    /// </summary>
    void SetImage(string filePath);
}

internal class SCImage : IImage
{
    private readonly SCPresentation presentation;
    private readonly StringValue picReference;
    private readonly OpenXmlPart openXmlPart;
    private byte[]? bytes;

    private SCImage(
        ImagePart imagePart,
        StringValue picReference,
        OpenXmlPart openXmlPart,
        SCPresentation presentation)
    {
        this.SDKImagePart = imagePart;
        this.picReference = picReference;
        this.openXmlPart = openXmlPart;

        this.presentation = presentation;
        this.MIME = this.SDKImagePart.ContentType;
    }

    public string MIME { get; }

    public Task<byte[]> BinaryData => this.GetBinaryData();

    public string Name => this.GetName();

    internal ImagePart SDKImagePart { get; private set; }

    public void SetImage(Stream stream)
    {
        var isSharedImagePart = this.presentation.ImageParts.Count(imgPart => imgPart == this.SDKImagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = RelatedIdGenerator.Generate();
            this.SDKImagePart = this.openXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.picReference.Value = rId;
        }

        stream.Position = 0;
        this.SDKImagePart.FeedData(stream);
        this.bytes = null; // resets cache
    }

    public void SetImage(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);

        this.SetImage(stream);
    }

    public void SetImage(string filePath)
    {
        byte[] sourceBytes = File.ReadAllBytes(filePath);
        this.SetImage(sourceBytes);
    }

    internal static SCImage ForPicture(Shape pictureShape, OpenXmlPart openXmlPart, StringValue picReference)
    {
        var imagePart = (ImagePart)openXmlPart.GetPartById(picReference.Value!);

        return new SCImage(imagePart, picReference, openXmlPart, pictureShape.SlideBase.PresentationInternal);
    }

    internal static SCImage? ForBackground(SCSlide slide)
    {
        var pBackground = slide.SDKSlidePart.Slide.CommonSlideData!.Background;
        if (pBackground == null)
        {
            return null;
        }

        var aBlipFill = pBackground.Descendants<A.BlipFill>().SingleOrDefault();
        var picReference = aBlipFill?.Blip?.Embed;
        if (picReference == null)
        {
            return null;
        }

        var imagePart = (ImagePart)slide.SDKSlidePart.GetPartById(picReference.Value!);
        var backgroundImage = new SCImage(imagePart, picReference, slide.SDKSlidePart, slide.PresentationInternal);

        return backgroundImage;
    }

    internal static SCImage? ForAutoShapeFill(Shape autoShape, TypedOpenXmlPart slidePart, A.BlipFill aBlipFill)
    {
        var picReference = aBlipFill.Blip?.Embed;
        if (picReference == null)
        {
            return null;
        }

        var imagePart = (ImagePart)slidePart.GetPartById(picReference.Value!);

        return new SCImage(imagePart, picReference, slidePart, autoShape.SlideBase.PresentationInternal);
    }

    internal static SCImage Create(ImagePart imagePart, MasterPicture masterPic, StringValue stringValue, SlideMasterPart sldMasterPart)
    {
        return new SCImage(imagePart, stringValue, sldMasterPart, masterPic.SlideBase.PresentationInternal);
    }

    internal static SCImage Create(ImagePart imagePart, LayoutPicture layoutPic, StringValue stringValue, SlideLayoutPart slideLayoutPart)
    {
        return new SCImage(imagePart, stringValue, slideLayoutPart, layoutPic.SlideBase.PresentationInternal);
    }

    private string GetName()
    {
        return Path.GetFileName(this.SDKImagePart.Uri.ToString());
    }

    private async Task<byte[]> GetBinaryData()
    {
        if (this.bytes != null)
        {
            return this.bytes; // return from cache
        }

        Stream stream = this.SDKImagePart.GetStream();
        this.bytes = new byte[stream.Length];
        await stream.ReadAsync(this.bytes, 0, (int)stream.Length).ConfigureAwait(false);
        stream.Close();
        return this.bytes;
    }
}