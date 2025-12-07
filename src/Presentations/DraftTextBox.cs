using System;
using GeometryEnum = ShapeCrawler.Geometry;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft text box.
/// </summary>
public sealed class DraftTextBox
{
    internal string TextBoxName { get; private set; } = "Text Box";

    internal int PosX { get; private set; }

    internal int PosY { get; private set; }

    internal int BoxWidth { get; private set; } = 100;

    internal int BoxHeight { get; private set; } = 50;

    internal string? Content { get; private set; }

    internal Color? HighlightColor { get; private set; }

    internal GeometryEnum ShapeGeometry { get; private set; } = GeometryEnum.Rectangle;

    /// <summary>
    ///     Sets the geometry type of the text box.
    /// </summary>
    public DraftTextBox Geometry(GeometryEnum geometry)
    {
        this.ShapeGeometry = geometry;
        return this;
    }

    /// <summary>
    ///     Sets text content.
    /// </summary>
    public DraftTextBox Text(string text)
    {
        this.Content = text;
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

    internal DraftTextBox NameMethod(string name)
    {
        this.TextBoxName = name;
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
}