using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft table shape builder.
/// </summary>
public sealed class DraftTableShape
{
    internal int ShapeX { get; private set; }

    internal int ShapeY { get; private set; }

    internal DraftTable? DraftTableBuilder { get; private set; }

    /// <summary>
    ///     Sets X-position in points.
    /// </summary>
    public DraftTableShape X(int x)
    {
        this.ShapeX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position in points.
    /// </summary>
    public DraftTableShape Y(int y)
    {
        this.ShapeY = y;
        return this;
    }

    /// <summary>
    ///     Configures the table using a nested builder.
    /// </summary>
    public DraftTableShape Table(Action<DraftTable> configure)
    {
        var builder = new DraftTable();
        configure(builder);
        this.DraftTableBuilder = builder;
        return this;
    }
}