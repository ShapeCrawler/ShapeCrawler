using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;

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
    ///     Gets or sets the slide name. Returns an empty string when no name is assigned.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether shapes from the master slide are shown.
    ///     The default is <see langword="true"/>.
    /// </summary>
    bool ShowMasterShapes { get; set; }

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
    ///     Includes the slide in the slide show, reversing <see cref="Hide"/>.
    /// </summary>
    void Show();

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
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
    void SaveImage(Stream stream);

    /// <summary>
    ///     Saves the slide image to the specified file.
    /// </summary>
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
    void SaveImage(string file);

    /// <summary>
    ///     Saves the slide as a PNG image with the specified pixel dimensions to the stream.
    ///     The slide is stretched to fill the image, including when its aspect ratio differs.
    /// </summary>
    /// <param name="stream">Destination stream, which remains open and is rewound when seekable.</param>
    /// <param name="width">Positive image width in pixels.</param>
    /// <param name="height">Positive image height in pixels.</param>
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
    void SaveImage(Stream stream, int width, int height);

    /// <summary>
    ///     Saves the slide as a PNG image with the specified pixel dimensions to the file.
    ///     The slide is stretched to fill the image, including when its aspect ratio differs.
    /// </summary>
    /// <param name="file">Destination file.</param>
    /// <param name="width">Positive image width in pixels.</param>
    /// <param name="height">Positive image height in pixels.</param>
    [Experimental("EXP001", Message = "This slide image generation API is experimental.")]
    void SaveImage(string file, int width, int height);

    /// <summary>
    ///     Gets the first shape in the slide.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T First<T>();
}
