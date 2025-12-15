namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft paragraph for fluent API.
/// </summary>
public sealed class DraftParagraph
{
    internal string? Content { get; private set; }

    internal bool IsBulletedList { get; private set; }

    internal string BulletCharacter { get; private set; } = "•";

    /// <summary>
    ///     Sets paragraph text.
    /// </summary>
    public DraftParagraph Text(string text)
    {
        this.Content = text;
        return this;
    }

    /// <summary>
    ///     Makes this paragraph a bulleted list item.
    /// </summary>
    public DraftParagraph BulletedList()
    {
        this.IsBulletedList = true;
        return this;
    }

    /// <summary>
    ///     Makes this paragraph a bulleted list item with a custom character.
    /// </summary>
    public DraftParagraph BulletedList(string character)
    {
        this.IsBulletedList = true;
        this.BulletCharacter = character;
        return this;
    }
}
