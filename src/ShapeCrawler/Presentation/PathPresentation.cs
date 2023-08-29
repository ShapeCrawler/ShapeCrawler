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

    public void Save() => this.presentation.Save(path);
    void IValidateable.Validate() => this.presentation.Validate();

    public void Copy(string newPath)
    {
        this.path = newPath;
        this.Save();
    }

    public void Copy(Stream stream) => this.presentation.Save(stream);
    public ISlideCollection Slides => this.presentation.Slides;

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
    public byte[] BinaryData => this.presentation.BinaryData;
    public ISectionCollection Sections => this.presentation.Sections;
    public IHeaderAndFooter HeaderAndFooter => this.presentation.HeaderAndFooter;
}