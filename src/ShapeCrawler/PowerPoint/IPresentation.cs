using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a presentation document.
/// </summary>
public interface IPresentation : IDisposable
{
    /// <summary>
    ///     Gets the presentation slides.
    /// </summary>
    ISlideCollection Slides { get; }

    /// <summary>
    ///     Gets the presentation slides width.
    /// </summary>
    int SlideWidth { get; }

    /// <summary>
    ///     Gets the presentation slides height.
    /// </summary>
    int SlideHeight { get; }

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
    PresentationDocument SDKPresentation { get; }

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();

    /// <summary>
    ///     Saves presentation in specified file path.
    /// </summary>
    void SaveAs(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void SaveAs(Stream stream);

    /// <summary>
    ///     Closes the presentation.
    /// </summary>
    void Close();
}