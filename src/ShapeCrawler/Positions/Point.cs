namespace ShapeCrawler.Positions;

/// <summary>
///     Represents a point in 2D space.
/// </summary>
public class Point
{
    internal Point(decimal x, decimal y)
    {
        this.X = (int)x;
        this.Y = (int)y;
    }
    
    internal int X { get; }
    
    internal int Y { get; }
}