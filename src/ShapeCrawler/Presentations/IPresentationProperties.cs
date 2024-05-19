namespace ShapeCrawler;

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
    ///     Gets collection of the slide masters.
    /// </summary>
    ISlideMasterCollection SlideMasters { get; }
    
    /// <summary>
    ///     Gets section collection.
    /// </summary>
    ISections Sections { get; }

    /// <summary>
    ///     Gets Header and Footer manager.
    /// </summary>
    IFooter Footer { get; }

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();
    
    /// <summary>
    ///     Gets a presentation byte array.
    /// </summary>
    byte[] AsByteArray();
}