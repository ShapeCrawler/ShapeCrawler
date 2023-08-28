using System.IO;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal sealed record PresentationCore
{
    private readonly SlideSize slideSize;

    internal PresentationCore(byte[] bytes)
        : this(new MemoryStream(bytes))
    {
    }

    internal PresentationCore(Stream stream)
    {
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new  Slides(this.sdkPresDocument.PresentationPart!.SlideParts);
        this.HeaderAndFooter = new HeaderAndFooter(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
    }

    private PresentationDocument sdkPresDocument { get; set; }

    public ISlideCollection Slides { get; }

    public int SlideHeight
    {
        get => this.slideSize.Height();
        set => this.slideSize.UpdateHeight(value);
    }

    public int SlideWidth
    {
        get => this.slideSize.Width();
        set => this.slideSize.UpdateWidth(value);
    }

    public ISlideMasterCollection SlideMasters { get; }

    public byte[] BinaryData => this.GetByteArray();

    public ISectionCollection Sections { get; }

    public IHeaderAndFooter HeaderAndFooter { get; }

    public void Save(string path)
    {
        var cloned = this.sdkPresDocument.Clone(path);
        cloned.Dispose();
    }

    public void Save(Stream stream)
    {
        this.sdkPresDocument.Clone(stream);
    }
    
    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.sdkPresDocument.Clone(stream);

        return stream.ToArray();
    }
}