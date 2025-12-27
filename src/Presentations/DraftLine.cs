using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft line.
/// </summary>
public sealed class DraftLine
{
    internal string DraftName { get; private set; } = "Line";

    internal int DraftX { get; private set; }

    internal int DraftY { get; private set; }

    internal int DraftWidth { get; private set; } = 100;

    internal int DraftHeight { get; private set; }

    internal DraftStroke? DraftStroke { get; private set; }

    /// <summary>
    ///     Sets name.
    /// </summary>
    public DraftLine Name(string name)
    {
        this.DraftName = name;
        return this;
    }

    /// <summary>
    ///     Sets X-position of the start point in points.
    /// </summary>
    public DraftLine X(int x)
    {
        this.DraftX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position of the start point in points.
    /// </summary>
    public DraftLine Y(int y)
    {
        this.DraftY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points (endX = startX + width).
    /// </summary>
    public DraftLine Width(int width)
    {
        this.DraftWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points (endY = startY + height).
    /// </summary>
    public DraftLine Height(int height)
    {
        this.DraftHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the line stroke.
    /// </summary>
    public DraftLine Line(Action<DraftStroke> configure)
    {
        this.DraftStroke = new DraftStroke();
        configure(this.DraftStroke);
        return this;
    }

    internal DocumentFormat.OpenXml.Drawing.LineEndValues? DraftHeadEndType { get; private set; }

    internal DocumentFormat.OpenXml.Drawing.LineEndValues? DraftTailEndType { get; private set; }

    /// <summary>
    ///     Sets the arrow head type for the end of the line.
    /// </summary>
    public DraftLine EndArrow(DocumentFormat.OpenXml.Drawing.LineEndValues type)
    {
        this.DraftTailEndType = type;
        return this;
    }

    /// <summary>
    ///     Sets the arrow head type for the start of the line.
    /// </summary>
    public DraftLine StartArrow(DocumentFormat.OpenXml.Drawing.LineEndValues type)
    {
        this.DraftHeadEndType = type;
        return this;
    }
}