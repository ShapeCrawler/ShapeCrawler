namespace ShapeCrawler;

/// <summary>
///     Represents a point in 2D space.
/// </summary>
public sealed class Point
{
    internal Point(decimal x, decimal y)
    {
        this.X = (int)x;
        this.Y = (int)y;
    }
    
    internal int X { get; }
    
    internal int Y { get; }
}