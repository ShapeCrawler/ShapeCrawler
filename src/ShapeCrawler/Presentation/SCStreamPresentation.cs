using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal sealed record SCStreamPresentation : IPresentationInternal
{
    private readonly PresentationCore presentationCore;
    private readonly Stream stream;

    internal SCStreamPresentation(Stream userStream)
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

    public void Save(string path)
    {
        this.presentationCore.Save(path);
    }

    public void Save(Stream newStream)
    {
        this.presentationCore.Save(stream);
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