// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

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
    
    /// <summary>
    ///     Gets the X coordinate.
    /// </summary>
    public int X { get; }

    /// <summary>
    ///     Gets the Y coordinate.
    /// </summary>
    public int Y { get; }
}