using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

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

internal sealed class SCImage : IImage
{
    private readonly SCPresentation presentation;
    private readonly StringValue blipEmbed;
    private readonly OpenXmlPart openXmlPart;
    private byte[]? bytes;

    private SCImage(
        ImagePart imagePart,
        StringValue blipEmbed,
        OpenXmlPart openXmlPart,
        SCPresentation presentation)
    {
        this.SDKImagePart = imagePart;
        this.blipEmbed = blipEmbed;
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
            this.blipEmbed.Value = rId;
        }

        stream.Position = 0;
        this.SDKImagePart.FeedData(stream);
        this.bytes = null; // to reset cache
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

    internal static SCImage ForPicture(SCShape pictureSCShape, OpenXmlPart openXmlPart, StringValue? blipEmbed)
    {
        var imagePart = (ImagePart)openXmlPart.GetPartById(blipEmbed!.Value!);

        return new SCImage(imagePart, blipEmbed, openXmlPart, pictureSCShape.SlideBase.PresentationInternal);
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

    internal static SCImage? ForAutoShapeFill(SlideStructure slideObject, TypedOpenXmlPart slidePart, A.BlipFill aBlipFill)
    {
        var picReference = aBlipFill.Blip?.Embed;
        if (picReference == null)
        {
            return null;
        }

        var imagePart = (ImagePart)slidePart.GetPartById(picReference.Value!);

        return new SCImage(imagePart, picReference, slidePart, slideObject.PresentationInternal);
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