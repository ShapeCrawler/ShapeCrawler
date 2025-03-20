using System.ComponentModel;

namespace ShapeCrawler;

/// <summary>
///     Represents a point in 2D space.
/// </summary>
public sealed record Point(
    [property: Description("Gets X-coordinate in points")]
    decimal X,
    [property: Description("Gets Y-coordinate in points")]
    decimal Y
    );