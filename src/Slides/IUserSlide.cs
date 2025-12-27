using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

#if DEBUG
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a regular PowerPoint Slide.
/// </summary>
public interface IUserSlide
{
    /// <summary>
    ///     Gets or sets custom data. Returns <see langword="null"/> if the custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets slide layout.
    /// </summary>
    ILayoutSlide LayoutSlide { get; }

    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }

    /// <summary>
    ///     Gets the shape collection.
    /// </summary>
    IUserSlideShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets the slide notes.
    /// </summary>
    ITextBox? Notes { get; }

    /// <summary>
    ///     Gets the slide fill.
    /// </summary>
    IShapeFill Fill { get; }

    /// <summary>
    ///     Gets all text content from shapes on the slide.
    /// </summary>
    public IList<ITextBox> GetTexts();

    /// <summary>
    ///     Hides slide.
    /// </summary>
    void Hide();

    /// <summary>
    ///     Gets a value indicating whether the slide is hidden.
    /// </summary>
    bool Hidden();

    /// <summary>
    ///     Adds specified lines to the slide notes.
    /// </summary>
    void AddNotes(IEnumerable<string> lines);

    /// <summary>
    ///     Gets element by name.
    /// </summary>
    /// <param name="name">element name.</param>
    IShape Shape(string name);

    /// <summary>
    ///     Gets element by ID.
    /// </summary>
    IShape Shape(int id);

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Shape<T>(string name)
        where T : IShape;

    /// <summary>
    ///     Removes the slide.
    /// </summary>
    void Remove();

    /// <summary>
    ///     Saves the slide image to the specified stream.
    /// </summary>
#if NET10_0_OR_GREATER
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
#else
    [Experimental("EXP001")]
#endif
    void SaveImageTo(Stream stream);

    /// <summary>
    ///     Saves the slide image to the specified file.
    /// </summary>
#if NET10_0_OR_GREATER
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
#else
    [Experimental("EXP001")]
#endif
    void SaveImageTo(string file);

    /// <summary>
    ///     Gets a copy of the underlying parent <see cref="PresentationPart"/>.
    /// </summary>
    PresentationPart GetSdkPresentationPart(); // NOSONAR

    /// <summary>
    ///     Gets the first shape in the slide.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T First<T>();
}