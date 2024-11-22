#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a point.
/// </summary>
public readonly ref struct Point
{
    internal Point(int x, int y)
    {
        this.X = x;
        this.Y = y;
    }

    internal Point(decimal x, decimal y)
    {
        this.X = (int)x;
        this.Y = (int)y;
    }

    /// <summary>
    ///     Gets the X coordinate.
    /// </summary>
    public int X { get; }

    /// <summary>
    ///     Gets the Y coordinate.
    /// </summary>
    public int Y { get; }
}