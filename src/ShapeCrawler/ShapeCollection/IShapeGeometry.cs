#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the complete geometry of a shape.
/// </summary>
public interface IShapeGeometry
{
    /// <summary>
    ///     Gets or sets the defined geometry form of the shape.
    /// </summary>
    Geometry GeometryType { get; set; }

    /// <summary>
    ///     Gets or sets the size of the corners in this shape. Range: 0.0-1.0.
    /// </summary>
    /// <remarks>
    ///     Applicable only to rounded rectangle geometry.
    ///     Value 0.0 makes shape appear as a regular rectangle with no corners.
    ///     Value 1.0 is the maximum size of a corner: 1/2 length of shortest side.
    /// </remarks>
    decimal? CornerSize { get; set; }
}