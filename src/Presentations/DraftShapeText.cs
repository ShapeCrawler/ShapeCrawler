using System;
using System.Collections.Generic;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft shape text for fluent API.
/// </summary>
public sealed class DraftShapeText
{
    private readonly List<DraftParagraph> paragraphs = [];

    /// <summary>
    ///     Gets the draft paragraphs.
    /// </summary>
    internal IReadOnlyList<DraftParagraph> Paragraphs => this.paragraphs;

    /// <summary>
    ///     Adds a paragraph with specified text.
    /// </summary>
    public DraftShapeText Paragraph(string text)
    {
        var draftParagraph = new DraftParagraph();
        draftParagraph.Text(text);
        this.paragraphs.Add(draftParagraph);
        return this;
    }

    /// <summary>
    ///     Adds a paragraph configured using a nested builder.
    /// </summary>
    public DraftShapeText Paragraph(Action<DraftParagraph> configure)
    {
        var draftParagraph = new DraftParagraph();
        configure(draftParagraph);
        this.paragraphs.Add(draftParagraph);
        return this;
    }
}