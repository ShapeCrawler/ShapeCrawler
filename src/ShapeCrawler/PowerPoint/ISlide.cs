using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

/// <summary>
///     Represents a slide.
/// </summary>
public interface ISlide : ISlideStructure
{
    /// <summary>
    ///     Gets background image if it exist, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets or sets custom data. It returns <see langword="null"/> if custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets a value indicating whether the slide is hidden.
    /// </summary>
    bool Hidden { get; }

    /// <summary>
    ///     Gets referenced Slide Layout.
    /// </summary>
    ISlideLayout SlideLayout { get; }

    /// <summary>
    ///     Gets instance of <see cref="DocumentFormat.OpenXml.Packaging.SlidePart"/> class of the underlying Open XML SDK.
    /// </summary>
    SlidePart SDKSlidePart { get; }

    /// <summary>
    /// Gets a list of all textboxes on that slide, including those in tables.
    /// </summary>
    public IList<ITextFrame> GetAllTextFrames();

    /// <summary>
    ///     Hides slide.
    /// </summary>
    void Hide();

#if DEBUG
    /// <summary>
    ///     Saves slide as PNG image.
    /// </summary>
    void SaveAsPng(System.IO.Stream stream);
#endif
}