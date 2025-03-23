using System.ComponentModel;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a point in 2D space.
/// </summary>
public sealed record Point( // TODO: should be struct?
    [property: Description("Gets X-coordinate in points")]
    decimal X,
    [property: Description("Gets Y-coordinate in points")]
    decimal Y);