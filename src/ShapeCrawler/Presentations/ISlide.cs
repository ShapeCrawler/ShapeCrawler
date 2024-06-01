using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
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
    ///     Gets underlying instance of <see cref="DocumentFormat.OpenXml.Packaging.SlidePart"/>.
    /// </summary>
    SlidePart SDKSlidePart { get; }
    
    /// <summary>
    ///     Gets shape collection.
    /// </summary>
    ISlideShapes Shapes { get; }

    /// <summary>
    ///     Gets slide notes as a single text frame.
    /// </summary>
    ITextFrame? Notes { get; }

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
    
    /// <summary>
    ///     Adds a notes slide if there isn't already one.
    /// </summary>
    void AddNotesIfEmpty();
    
#if DEBUG
    
    /// <summary>
    ///     Saves slide as PNG image.
    /// </summary>
    void SaveAsPng(System.IO.Stream stream);
#endif
}