namespace ShapeCrawler.SlideMasters;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets parent Slide Master.
    /// </summary>
    ISlideMaster SlideMaster { get; }

    /// <summary>
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }
}