using System.IO;

namespace ShapeCrawler;

internal sealed record StreamPresentation : IValidateable
{
    private readonly Presentation presentation;
    private readonly Stream userStream;

    internal StreamPresentation(Stream userStream)
    {
        this.userStream = userStream;
        var internalStream = new MemoryStream();
        this.userStream.Position = 0;
        userStream.CopyTo(internalStream);
        this.presentation = new Presentation(internalStream);
    }

    public void Save() => this.presentation.CopyTo(this.userStream);
    void IValidateable.Validate() => this.presentation.Validate();
    public void CopyTo(string path) => this.presentation.CopyTo(path);
    public void CopyTo(Stream stream) => this.presentation.CopyTo(stream);
    public ISlides Slides => this.presentation.Slides;
    public int SlideWidth
    {
        get => this.presentation.SlideWidth;
        set => this.presentation.SlideWidth = value;
    }
    public int SlideHeight
    {
        get => this.presentation.SlideHeight;
        set => this.presentation.SlideHeight = value;
    }
    public ISlideMasterCollection SlideMasters => this.presentation.SlideMasters;
    public byte[] AsByteArray() => this.presentation.AsByteArray();
    public ISections Sections => this.presentation.Sections;
    public IHeaderAndFooter HeaderAndFooter => this.presentation.HeaderAndFooter;
}