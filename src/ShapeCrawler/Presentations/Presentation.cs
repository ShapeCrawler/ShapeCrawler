using System.IO;
using System.Reflection;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shared;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <inheritdoc cref="IPresentation"/>
public sealed class Presentation : IPresentation
{
    private IValidateable validateable;

    /// <summary>
    ///     Opens existing presentation from specified path.
    /// </summary>
    public Presentation(string path)
    {
        this.validateable = new PathPresentation(path);
    }

    /// <summary>
    ///     Opens existing presentation from specified stream.
    /// </summary>
    public Presentation(Stream stream)
    {
        this.validateable = new StreamPresentation(stream);
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public Presentation()
    {
        var assets = new Assets(Assembly.GetExecutingAssembly());
        var stream = assets.StreamOf("new-presentation.pptx");
        this.validateable = new StreamPresentation(stream);
        this.validateable.FileProperties.Modified =
            this.validateable.FileProperties.Created = SCSettings.TimeProvider.UtcNow;
    }

    /// <inheritdoc />
    public ISlides Slides => this.validateable.Slides;

    /// <inheritdoc />
    public decimal SlideWidth
    {
        get => this.validateable.SlideWidth; 
        set => this.validateable.SlideWidth = value;
    }
    
    /// <inheritdoc />
    public decimal SlideHeight
    {
        get => this.validateable.SlideHeight;
        set => this.validateable.SlideHeight = value;
    }
    
    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters => this.validateable.SlideMasters;
    
    /// <inheritdoc />
    public ISections Sections => this.validateable.Sections;

    /// <inheritdoc />
    public IFooter Footer => this.validateable.Footer;
    
    /// <inheritdoc />
    public IFileProperties FileProperties => this.validateable.FileProperties;

    /// <inheritdoc />
    public void Save() => this.validateable.Save();
    
    /// <inheritdoc />
    public ISlide Slide(int number) => this.validateable.Slide(number);
    
    /// <summary>
    ///     Gets a presentation byte array.
    /// </summary>
    public byte[] AsByteArray() => this.validateable.AsByteArray();

    /// <inheritdoc />
    public void SaveAs(string path)
    {
        this.validateable.CopyTo(path);
        this.validateable = new PathPresentation(path);
    }

    /// <inheritdoc />
    public void SaveAs(Stream stream)
    {
        this.validateable.CopyTo(stream);
        this.validateable = new StreamPresentation(stream);
    }
    
    /// <summary>
    ///     Gets Slide Master by number.
    /// </summary>
    public ISlideMaster SlideMaster(int number) => this.SlideMasters[number - 1];

    internal void Validate() => this.validateable.Validate();
}