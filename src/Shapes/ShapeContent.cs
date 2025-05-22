// ReSharper disable once CheckNamespace
// ReSharper disable InconsistentNaming
#pragma warning disable IDE0130

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Main shape content.
/// </summary>
public enum ShapeContent
{
    /// <summary>
    ///     Shape.
    /// </summary>
    Shape,
    
    /// <summary>
    ///     Audio.
    /// </summary>
    Audio,

    /// <summary>
    ///     Chart.
    /// </summary>
    Chart,

    /// <summary>
    ///     Line.
    /// </summary>
    Line,

    /// <summary>
    ///     Grouped elements.
    /// </summary>
    GroupedElement,

    /// <summary>
    ///     OLE Object.
    /// </summary>
    OLEObject,

    /// <summary>
    ///     Image.
    /// </summary>
    Image,

    /// <summary>
    ///     Video.
    /// </summary>
    Video,

    /// <summary>
    ///     Table.
    /// </summary>
    Table
}