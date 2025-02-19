namespace ShapeCrawler.Shapes;

/// <summary>
///     Root (non-grouped) Auto Shape.
/// </summary>
internal interface IRootShape : IShape
{
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    void Duplicate();
}