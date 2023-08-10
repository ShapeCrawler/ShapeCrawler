namespace ShapeCrawler;

/// <summary>
///     Represents slide object.
/// </summary>
public interface ISlideStructure
{
    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }
    
    /// <summary>
    ///     Gets collection of shapes.
    /// </summary>
    ISlideShapeCollection Shapes { get; }
}