using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal sealed record StreamPresentation : ICopyablePresentation
{
    private readonly PresentationCore presentationCore;
    private readonly Stream stream;

    internal StreamPresentation(Stream userStream)
    {
        this.stream = userStream;
        var internalStream = new MemoryStream();
        userStream.CopyTo(internalStream);
        this.presentationCore = new PresentationCore(internalStream);
    }

    public void Save()
    {
        this.presentationCore.Save(this.stream);
    }

    public void Copy(string path)
    {
        this.presentationCore.Save(path);
    }

    public void Copy(Stream userStream)
    {
        this.presentationCore.Save(userStream);
    }

    /// <inheritdoc />
    public ISlideCollection Slides => this.presentationCore.Slides;

    /// <inheritdoc />
    public int SlideWidth
    {
        get => this.presentationCore.SlideWidth;
        set => this.presentationCore.SlideWidth = value;
    }

    /// <inheritdoc />
    public int SlideHeight
    {
        get => this.presentationCore.SlideHeight;
        set => this.presentationCore.SlideHeight = value;
    }

    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters => this.presentationCore.SlideMasters;

    /// <inheritdoc />
    public byte[] BinaryData => this.presentationCore.BinaryData;

    /// <inheritdoc />
    public ISectionCollection Sections => this.presentationCore.Sections;

    /// <inheritdoc />
    public PresentationDocument SDKPresentationDocument => this.presentationCore.SDKPresentationDocument;

    /// <inheritdoc />
    public IHeaderAndFooter HeaderAndFooter => this.presentationCore.HeaderAndFooter;
}