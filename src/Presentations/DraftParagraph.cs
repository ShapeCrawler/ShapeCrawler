using System;

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
    ///    Gets draft font.
    /// </summary>
    internal DraftFont? FontDraft { get; private set; }

    /// <summary>
    ///     Sets paragraph text.
    /// </summary>
    public DraftParagraph Text(string text)
    {
        this.Content = text;
        return this;
    }

    /// <summary>
    ///     Configures font using a nested builder.
    /// </summary>
    public DraftParagraph Font(Action<DraftFont> configure)
    {
        this.FontDraft = new DraftFont();
        configure(this.FontDraft);
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
