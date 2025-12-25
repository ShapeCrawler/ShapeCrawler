using System;
using System.Collections.Generic;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft text box.
/// </summary>
public sealed class DraftTextBox
{
    /// <summary>
    ///    Gets or sets a value indicating whether this is a text box.
    /// </summary>
    internal bool IsTextBox { get; set; }

    internal string TextBoxName { get; private set; } = "Text Box";

    internal int PosX { get; private set; }

    internal int PosY { get; private set; }

    internal int BoxWidth { get; private set; } = 100;

    internal int BoxHeight { get; private set; } = 50;

    internal string? Content { get; private set; }

    internal Color? HighlightColor { get; private set; }

    internal Geometry ShapeGeometry { get; private set; } = ShapeCrawler.Geometry.Rectangle;

    internal List<DraftParagraph> Paragraphs { get; } = [];

    /// <summary>
    ///    Gets draft font.
    /// </summary>
    internal DraftFont? FontDraft { get; private set; }

    /// <summary>
    ///     Sets the geometry type of the text box.
    /// </summary>
    public DraftTextBox Geometry(Geometry geometry)
    {
        this.ShapeGeometry = geometry;
        return this;
    }

    /// <summary>
    ///     Configures font using a nested builder.
    /// </summary>
    public DraftTextBox Font(Action<DraftFont> configure)
    {
        this.FontDraft = new DraftFont();
        configure(this.FontDraft);
        return this;
    }

    /// <summary>
    ///     Sets text content.
    /// </summary>
    public DraftTextBox TextBox(string text)
    {
        this.Content = text;
        return this;
    }

    /// <summary>
    ///     Configures shape text using a nested builder.
    /// </summary>
    public DraftTextBox TextBox(Action<DraftShapeText> configure)
    {
        var draftShapeText = new DraftShapeText();
        configure(draftShapeText);

        this.Content = null;
        this.Paragraphs.Clear();
        this.Paragraphs.AddRange(draftShapeText.Paragraphs);
        return this;
    }

    /// <summary>
    ///     Sets text highlight color.
    /// </summary>
    public DraftTextBox TextHighlightColor(Color color)
    {
        this.HighlightColor = color;
        return this;
    }

    /// <summary>
    ///     Sets name.
    /// </summary>
    public DraftTextBox Name(string name) => this.NameMethod(name);

    /// <summary>
    ///     Sets X-position.
    /// </summary>
    public DraftTextBox X(int x)
    {
        this.PosX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position.
    /// </summary>
    public DraftTextBox Y(int y)
    {
        this.PosY = y;
        return this;
    }

    /// <summary>
    ///     Sets width.
    /// </summary>
    public DraftTextBox Width(int width)
    {
        this.BoxWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height.
    /// </summary>
    public DraftTextBox Height(int height)
    {
        this.BoxHeight = height;
        return this;
    }

    /// <summary>
    ///     Adds paragraph.
    /// </summary>
    public DraftTextBox Paragraph(string content)
    {
        this.Content = AppendParagraph(this.Content, content);
        return this;
    }

    /// <summary>
    ///     Configures a paragraph using a nested builder.
    /// </summary>
    public DraftTextBox Paragraph(Action<DraftParagraph> configure)
    {
        var draftParagraph = new DraftParagraph();
        configure(draftParagraph);
        this.Paragraphs.Add(draftParagraph);
        return this;
    }

    private static string AppendParagraph(string? current, string next)
    {
        if (string.IsNullOrEmpty(current))
        {
            return next;
        }

        return current + Environment.NewLine + next;
    }

    private DraftTextBox NameMethod(string name)
    {
        this.TextBoxName = name;
        return this;
    }
}