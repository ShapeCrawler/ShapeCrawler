namespace ShapeCrawler.SlideShape;

/// <summary>
///     Root (non-grouped) Auto Shape.
/// </summary>
internal interface IDuplicateableShape : IShape
{     
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    void Duplicate();
}