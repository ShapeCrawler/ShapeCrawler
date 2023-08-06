using System.IO;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a presentation document.
/// </summary>
internal interface IPresentationInternal
{
    /// <summary>
    ///     Gets the presentation slides.
    /// </summary>
    ISlideCollection Slides { get; }

    /// <summary>
    ///     Gets or sets presentation slides width in pixels.
    /// </summary>
    int SlideWidth { get; set; }

    /// <summary>
    ///     Gets or sets the presentation slides height.
    /// </summary>
    int SlideHeight { get; set; }

    /// <summary>
    ///     Gets collection of the slide masters.
    /// </summary>
    ISlideMasterCollection SlideMasters { get; }

    /// <summary>
    ///     Gets a presentation byte array.
    /// </summary>
    byte[] BinaryData { get; }

    /// <summary>
    ///     Gets section collection.
    /// </summary>
    ISectionCollection Sections { get; }

    /// <summary>
    ///     Gets copy of instance of <see cref="DocumentFormat.OpenXml.Packaging.PresentationDocument"/> class.
    /// </summary>
    PresentationDocument SDKPresentationDocument { get; }

    /// <summary>
    ///     Gets Header and Footer manager.
    /// </summary>
    IHeaderAndFooter HeaderAndFooter { get; }

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();

    /// <summary>
    ///     Saves presentation in specified file path.
    /// </summary>
    void Save(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void Save(Stream stream);
}