using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft rectangle shape.
/// </summary>
public sealed class DraftRectangle
{
    internal string DraftName { get; private set; } = "Rectangle";

    internal int DraftX { get; private set; }

    internal int DraftY { get; private set; }

    internal int DraftWidth { get; private set; } = 100;

    internal int DraftHeight { get; private set; } = 50;

    internal DraftSolidFill? SolidFillDraft { get; private set; }

    /// <summary>
    ///     Sets name.
    /// </summary>
    public DraftRectangle Name(string name)
    {
        this.DraftName = name;
        return this;
    }

    /// <summary>
    ///     Sets X-position.
    /// </summary>
    public DraftRectangle X(int x)
    {
        this.DraftX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position.
    /// </summary>
    public DraftRectangle Y(int y)
    {
        this.DraftY = y;
        return this;
    }

    /// <summary>
    ///     Sets width.
    /// </summary>
    public DraftRectangle Width(int width)
    {
        this.DraftWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height.
    /// </summary>
    public DraftRectangle Height(int height)
    {
        this.DraftHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the rectangle solid fill using a nested builder.
    /// </summary>
    public DraftRectangle SolidFill(Action<DraftSolidFill> configure)
    {
        this.SolidFillDraft = new DraftSolidFill();
        configure(this.SolidFillDraft);
        return this;
    }
}