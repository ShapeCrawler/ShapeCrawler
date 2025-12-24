namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft font.
/// </summary>
public sealed class DraftFont
{
    /// <summary>
    ///    Gets font size.
    /// </summary>
    internal decimal? SizeValue { get; private set; }

    /// <summary>
    ///    Gets a value indicating whether font is bold.
    /// </summary>
    internal bool IsBoldValue { get; private set; }

    /// <summary>
    ///     Sets font size.
    /// </summary>
    public DraftFont Size(int size)
    {
        this.SizeValue = size;
        return this;
    }

    /// <summary>
    ///     Sets font to bold.
    /// </summary>
    public DraftFont Bold()
    {
        this.IsBoldValue = true;
        return this;
    }
}