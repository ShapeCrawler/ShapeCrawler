#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the complete geometry of a shape.
/// </summary>
public interface IShapeGeometry
{
    /// <summary>
    ///     Gets or sets the geometry form.
    /// </summary>
    Geometry GeometryType { get; set; }

    /// <summary>
    ///     Gets or sets the corner size in percentages.
    /// </summary>
    /// <remarks>
    ///     Applicable only to rounded rectangle and top corners rounded geometry.
    ///     Value 0 makes shape appear as a regular rectangle with no corners.
    ///     Value 100 is the maximum size of a corner: 1/2 length of shortest side.
    /// </remarks>
    decimal CornerSize { get; set; }

    /// <summary>
    ///     Gets or sets the geometry adjustments.
    /// </summary>
    /// <remarks>
    ///     These are a set of geometry-specific adjustments with varying ranges and
    ///     meanings, unique to each kind of geometry.
    ///
    ///     An empty result may mean that a shape does not allow adjustments, or it
    ///     may mean that the user simply hasn't set any, so the adjustments are in
    ///     their default position(s).
    ///
    ///     Setting adjustments is only allowed for rounded or snipped rectangles,
    ///     in which case it adjusts the percentage of roundedness or snipping (0-100).
    /// </remarks>
    decimal[] Adjustments { get; set; }
}