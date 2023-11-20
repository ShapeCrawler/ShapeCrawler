using System.Collections.Generic;
#if DEBUG
using System.Threading.Tasks;
#endif

namespace ShapeCrawler;

/// <summary>
///     Represents a slide.
/// </summary>
public interface ISlide
{
    /// <summary>
    ///     Gets background image.
    /// </summary>
    IImage Background { get; }

    /// <summary>
    ///     Gets or sets custom data. It returns <see langword="null"/> if custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }
    
    /// <summary>
    ///     Gets referenced Slide Layout.
    /// </summary>
    ISlideLayout SlideLayout { get; }
    
    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }
    
    /// <summary>
    ///     Gets shape collection.
    /// </summary>
    ISlideShapes Shapes { get; }

    /// <summary>
    ///     List of all text frames on that slide.
    /// </summary>
    public IList<ITextFrame> TextFrames();

    /// <summary>
    ///     Hides slide.
    /// </summary>
    void Hide();
    
    /// <summary>
    ///     Gets a value indicating whether the slide is hidden.
    /// </summary>
    bool Hidden();
    
    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    IShape ShapeWithName(string autoShape);
    
    /// <summary>
    ///     Gets table by name.
    /// </summary>
    ITable TableWithName(string table);

#if DEBUG
    Task<string> ToHtml();
    
    /// <summary>
    ///     Saves slide as PNG image.
    /// </summary>
    void SaveAsPng(System.IO.Stream stream);
#endif
}