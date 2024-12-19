#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a presentation properties.
/// </summary>
public interface IPresentationProperties
{
    /// <summary>
    ///     Gets the presentation slides.
    /// </summary>
    ISlides Slides { get; }

    /// <summary>
    ///     Gets or sets presentation slides width in pixels.
    /// </summary>
    decimal SlideWidth { get; set; }

    /// <summary>
    ///     Gets or sets the presentation slides height.
    /// </summary>
    decimal SlideHeight { get; set; }

    /// <summary>
    ///     Gets the collection of the slide masters.
    /// </summary>
    ISlideMasterCollection SlideMasters { get; }
    
    /// <summary>
    ///     Gets the collection of sections.
    /// </summary>
    ISections Sections { get; }
    
    /// <summary>
    ///     Gets Header and Footer manager.
    /// </summary>
    IFooter Footer { get; }

    /// <summary>
    ///     Gets the proeprties of the file.
    /// </summary>
    /// <remarks>
    ///     These properties are not presentation-specific.
    /// </remarks>
    IFileProperties FileProperties { get; }
    
    /// <summary>
    ///     Returns slide with specified number.
    /// </summary>
    ISlide Slide(int number);

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();
    
    /// <summary>
    ///     Gets a presentation byte array.
    /// </summary>
    byte[] AsByteArray();
}