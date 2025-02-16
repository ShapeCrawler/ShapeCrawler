// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Shape type.
/// </summary>
public enum ShapeType
{
    /// <summary>
    ///     Shape.
    /// </summary>
    AutoShape,
    
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
    ///     Group.
    /// </summary>
    Group,

    /// <summary>
    ///     OLE Object.
    /// </summary>
    OleObject,

    /// <summary>
    ///     Picture.
    /// </summary>
    Picture,

    /// <summary>
    ///     Video.
    /// </summary>
    Video,

    /// <summary>
    ///     Table.
    /// </summary>
    Table
}