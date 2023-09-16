using System.IO;

namespace ShapeCrawler;

internal sealed record PathPresentation : IValidateable
{
    private readonly Presentation presentation;
    private string path;

    internal PathPresentation(string path)
    {
        this.path = path;
        this.presentation = new Presentation(File.ReadAllBytes(this.path));
    }

    public void Save() => this.presentation.CopyTo(path);
    void IValidateable.Validate() => this.presentation.Validate();

    public void CopyTo(string newPath)
    {
        this.path = newPath;
        this.Save();
    }

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