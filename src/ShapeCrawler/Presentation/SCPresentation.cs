using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

/// <inheritdoc cref="IPresentation"/>
public sealed record SCPresentation : IPresentation
{
    private ICopyablePresentation presentation;

    /// <summary>
    ///     Creates presentation from specified file path.
    /// </summary>
    public SCPresentation(string path)
    {
        this.presentation = new PathPresentation(path);
    }

    /// <summary>
    ///     Creates presentation from specified stream.
    /// </summary>
    public SCPresentation(Stream stream)
    {
        this.presentation = new StreamPresentation(stream);
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public SCPresentation()
    {
        this.presentation = new StreamPresentation(new MemoryStream());
    }

    /// <inheritdoc />
    public ISlideCollection Slides => this.presentation.Slides;

    /// <inheritdoc />
    public int SlideWidth
    {
        get => this.presentation.SlideWidth; 
        set => this.presentation.SlideWidth = value;
    }

    /// <inheritdoc />
    public int SlideHeight
    {
        get => this.presentation.SlideHeight;
        set => this.presentation.SlideHeight = value;
    }
    
    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters => this.presentation.SlideMasters;
    
    /// <inheritdoc />
    public byte[] BinaryData => this.presentation.BinaryData;
    
    /// <inheritdoc />
    public ISectionCollection Sections => this.presentation.Sections;
    
    /// <inheritdoc />
    public PresentationDocument SDKPresentationDocument => this.presentation.SDKPresentationDocument;
    
    /// <inheritdoc />
    public IHeaderAndFooter HeaderAndFooter => this.presentation.HeaderAndFooter;
    
    /// <inheritdoc />
    public void Save()
    {
        this.presentation.Save();
    }

    /// <inheritdoc />
    public void SaveAs(string path)
    {
        this.presentation.Copy(path);
        this.presentation = new PathPresentation(path);
    }

    /// <inheritdoc />
    public void SaveAs(Stream stream)
    {
        this.presentation.Copy(stream);
        this.presentation = new StreamPresentation(stream);
    }
}