using System.IO;
using System.Reflection;
using ShapeCrawler.Shared;

namespace ShapeCrawler;

/// <inheritdoc cref="IPresentation"/>
public sealed record SCPresentation : IPresentation
{
    private IValidateable validateable;
    
    /// <summary>
    ///     Creates presentation from specified file path.
    /// </summary>
    public SCPresentation(string path)
    {
        this.validateable = new PathPresentation(path);
    }

    /// <summary>
    ///     Creates presentation from specified stream.
    /// </summary>
    public SCPresentation(Stream stream)
    {
        this.validateable = new StreamPresentation(stream);
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public SCPresentation()
    {
        var assets = new Assets(Assembly.GetExecutingAssembly());
        var mStream = assets.StreamOf("new-presentation.pptx");
        this.validateable = new StreamPresentation(mStream);
    }

    /// <inheritdoc />
    public ISlideCollection Slides => this.validateable.Slides;

    /// <inheritdoc />
    public int SlideWidth
    {
        get => this.validateable.SlideWidth; 
        set => this.validateable.SlideWidth = value;
    }

    /// <inheritdoc />
    public int SlideHeight
    {
        get => this.validateable.SlideHeight;
        set => this.validateable.SlideHeight = value;
    }
    
    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters => this.validateable.SlideMasters;
    
    /// <inheritdoc />
    public byte[] BinaryData => this.validateable.BinaryData;
    
    /// <inheritdoc />
    public ISectionCollection Sections => this.validateable.Sections;
   
    /// <inheritdoc />
    public IHeaderAndFooter HeaderAndFooter => this.validateable.HeaderAndFooter;
    
    /// <inheritdoc />
    public void Save() => this.validateable.Save();

    /// <inheritdoc />
    public void SaveAs(string path)
    {
        this.validateable.Copy(path);
        this.validateable = new PathPresentation(path);
    }

    /// <inheritdoc />
    public void SaveAs(Stream stream)
    {
        this.validateable.Copy(stream);
        this.validateable = new StreamPresentation(stream);
    }

    internal void Validate()
    {
        this.validateable.Validate();
    }
}