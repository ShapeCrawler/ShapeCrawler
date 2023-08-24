﻿using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

/// <summary>
///     Represents a presentation properties.
/// </summary>
public interface IPresentationProperties
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
}