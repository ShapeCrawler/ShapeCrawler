using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal sealed record SCPathPresentation : ISavePresentation
{
    private readonly PresentationCore presentationCore;
    private string path;

    internal SCPathPresentation(string path)
    {
        this.path = path;
        this.presentationCore = new PresentationCore(File.ReadAllBytes(this.path));
    }

    public void Save()
    {
        this.presentationCore.Save(path);
    }

    public void Save(string newPath)
    {
        this.path = newPath;
        this.Save();
    }

    public void Save(Stream stream)
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