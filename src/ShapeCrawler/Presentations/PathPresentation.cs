using System.IO;

namespace ShapeCrawler.Presentations;

internal sealed record PathPresentation : IValidateable
{
    private readonly PresentationCore presentationCore;
    private string path;

    internal PathPresentation(string path)
    {
        this.path = path;
        this.presentationCore = new PresentationCore(File.ReadAllBytes(this.path));
    }

    public void Save() => this.presentationCore.CopyTo(this.path);
    void IValidateable.Validate() => this.presentationCore.Validate();

    public void CopyTo(string path)
    {
        this.path = path;
        this.Save();
    }

    public void CopyTo(Stream stream) => this.presentationCore.CopyTo(stream);
    public ISlides Slides => this.presentationCore.Slides;

    public decimal SlideWidth
    {
        get => this.presentationCore.SlideWidth;
        set => this.presentationCore.SlideWidth = value;
    }

    public decimal SlideHeight
    {
        get => this.presentationCore.SlideHeight;
        set => this.presentationCore.SlideHeight = value;
    }

    public ISlideMasterCollection SlideMasters => this.presentationCore.SlideMasters;
    public byte[] AsByteArray() => this.presentationCore.AsByteArray();
    public ISections Sections => this.presentationCore.Sections;
    public ISlide Slide(int number) => this.presentationCore.Slides[number - 1];

    public IFooter Footer => this.presentationCore.Footer;

    public IFileProperties FileProperties => this.presentationCore.FileProperties;
}
