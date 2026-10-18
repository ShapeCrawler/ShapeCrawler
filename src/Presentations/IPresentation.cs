using System;
using System.IO;
using System.Text.RegularExpressions;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a presentation.
/// </summary>
public interface IPresentation : IDisposable
{
    /// <summary>
    ///     Gets the presentation slides.
    /// </summary>
    ISlideCollection Slides { get; }

    /// <summary>
    ///     Gets or sets presentation slides width in points.
    /// </summary>
    decimal SlideWidth { get; set; }

    /// <summary>
    ///     Gets or sets the presentation slides height in points.
    /// </summary>
    decimal SlideHeight { get; set; }

    /// <summary>
    ///     Gets the collection of the slide masters.
    /// </summary>
    IMasterSlideCollection MasterSlides { get; }

    /// <summary>
    ///     Gets the collection of sections.
    /// </summary>
    ISectionCollection Sections { get; }

    /// <summary>
    ///     Gets footer.
    /// </summary>
    IFooter Footer { get; }

    /// <summary>
    ///     Gets presentation properties.
    /// </summary>
    IPresentationProperties Properties { get; }

    /// <summary>
    ///     Gets slide by number.
    /// </summary>
    IUserSlide Slide(int number);

    /// <summary>
    ///     Replaces literal text in slide text boxes, groups, and table cells and returns the match count.
    ///     Uses ordinal matching, case-sensitive by default, with the formatting rules of ITextBox.ReplaceText.
    ///     Does not search notes, masters, layouts, chart data, or SmartArt data parts.
    /// </summary>
    int ReplaceText(string oldText, string newText, bool ignoreCase = false);

    /// <summary>
    ///     Highlights literal matches in slide text boxes, groups, and table cells and returns their count.
    ///     Uses ITextBox.HighlightText formatting and matching rules. Excludes notes, masters, layouts, charts, and SmartArt data.
    /// </summary>
    int HighlightText(string text, Color color, bool ignoreCase = false);

    /// <summary>
    ///     Highlights nonempty regex matches in the same scope as HighlightText and returns their count.
    ///     Uses the supplied regex options/timeout. Earlier text boxes may already be updated if a later search throws.
    /// </summary>
    int HighlightRegex(Regex regex, Color color);

    /// <summary>
    ///     Replaces exact, case-sensitive DrawingML font family references throughout the presentation,
    ///     including themes, masters, layouts, notes, charts, and SmartArt. Other formatting is preserved.
    ///     Theme references such as +mn-lt continue to inherit from the updated theme.
    ///     Does not change embedded font binaries, embedded documents, or text in images.
    /// </summary>
    /// <param name="oldName">Nonempty source font family name.</param>
    /// <param name="newName">Nonempty replacement font family name.</param>
    void ReplaceFont(string oldName, string newName);

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();

    /// <summary>
    ///     Saves presentation and copies in the specified stream.
    /// </summary>
    void Save(Stream stream);

    /// <summary>
    ///     Saves presentation and copies in the specified file.
    /// </summary>
    public void Save(string file);

    /// <summary>
    ///     Saves a password-encrypted presentation to the specified stream, leaving it open.
    ///     Uses Office Agile encryption with AES-256 and SHA-512.
    /// </summary>
    /// <param name="stream">The writable destination stream.</param>
    /// <param name="password">The password required to open the entire presentation (1 to 255 characters).</param>
    void Save(Stream stream, string password);

    /// <summary>
    ///     Saves a password-encrypted presentation to the specified file.
    ///     Uses Office Agile encryption with AES-256 and SHA-512.
    /// </summary>
    /// <param name="file">The destination file.</param>
    /// <param name="password">The password required to open the entire presentation (1 to 255 characters).</param>
    void Save(string file, string password);

    /// <summary>
    ///     Returns the presentation as a string in Markdown format.
    /// </summary>
    public string AsMarkdown();

    /// <summary>
    ///     Returns the presentation as a base64 string.
    /// </summary>
    string AsBase64();
}
